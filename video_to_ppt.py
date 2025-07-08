#!/usr/bin/env python3
"""
Video to PowerPoint Conversion Pipeline
Converts video files into PowerPoint presentations with key frames and extracted content.
"""

import os
import cv2
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import speech_recognition as sr
import moviepy.editor as mp
from sklearn.cluster import KMeans
import argparse
import logging
from pathlib import Path
import json
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class VideoToPPTConverter:
    """Main class for converting video to PowerPoint presentation."""
    
    def __init__(self, video_path, output_path, config=None):
        self.video_path = Path(video_path)
        self.output_path = Path(output_path)
        self.config = config or self._default_config()
        self.extracted_frames = []
        self.transcription = ""
        self.frame_timestamps = []
        
    def _default_config(self):
        """Default configuration for the converter."""
        return {
            'max_frames': 10,
            'similarity_threshold': 0.8,
            'audio_extraction': True,
            'frame_analysis': True,
            'slide_layout': 'image_with_text',
            'font_size': 16,
            'title_font_size': 24
        }
    
    def extract_key_frames(self):
        """Extract key frames from video using similarity analysis."""
        logger.info("Extracting key frames from video...")
        
        cap = cv2.VideoCapture(str(self.video_path))
        frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        fps = cap.get(cv2.CAP_PROP_FPS)
        
        frames = []
        frame_histograms = []
        timestamps = []
        
        # Sample frames at regular intervals
        sample_interval = max(1, frame_count // (self.config['max_frames'] * 3))
        
        for i in range(0, frame_count, sample_interval):
            cap.set(cv2.CAP_PROP_POS_FRAMES, i)
            ret, frame = cap.read()
            if not ret:
                break
                
            # Calculate histogram for similarity comparison
            hist = cv2.calcHist([frame], [0, 1, 2], None, [50, 50, 50], [0, 256, 0, 256, 0, 256])
            frames.append(frame)
            frame_histograms.append(hist.flatten())
            timestamps.append(i / fps)
        
        cap.release()
        
        # Use K-means clustering to find most representative frames
        if len(frames) > self.config['max_frames']:
            kmeans = KMeans(n_clusters=self.config['max_frames'], random_state=42)
            clusters = kmeans.fit_predict(frame_histograms)
            
            # Select frame closest to each cluster center
            selected_indices = []
            for i in range(self.config['max_frames']):
                cluster_indices = np.where(clusters == i)[0]
                if len(cluster_indices) > 0:
                    cluster_center = kmeans.cluster_centers_[i]
                    distances = [np.linalg.norm(frame_histograms[idx] - cluster_center) 
                               for idx in cluster_indices]
                    closest_idx = cluster_indices[np.argmin(distances)]
                    selected_indices.append(closest_idx)
            
            self.extracted_frames = [frames[i] for i in sorted(selected_indices)]
            self.frame_timestamps = [timestamps[i] for i in sorted(selected_indices)]
        else:
            self.extracted_frames = frames
            self.frame_timestamps = timestamps
        
        logger.info(f"Extracted {len(self.extracted_frames)} key frames")
        return self.extracted_frames
    
    def extract_audio_and_transcribe(self):
        """Extract audio from video and transcribe to text."""
        if not self.config['audio_extraction']:
            return ""
        
        logger.info("Extracting audio and transcribing...")
        
        try:
            # Extract audio using moviepy
            video = mp.VideoFileClip(str(self.video_path))
            audio_path = self.output_path.parent / f"{self.video_path.stem}_audio.wav"
            video.audio.write_audiofile(str(audio_path), verbose=False, logger=None)
            
            # Transcribe audio
            recognizer = sr.Recognizer()
            with sr.AudioFile(str(audio_path)) as source:
                audio_data = recognizer.record(source)
                try:
                    self.transcription = recognizer.recognize_google(audio_data)
                    logger.info("Audio transcription completed")
                except sr.UnknownValueError:
                    logger.warning("Could not understand audio")
                    self.transcription = "Audio transcription not available"
                except sr.RequestError as e:
                    logger.error(f"Could not request results from Google Speech Recognition service; {e}")
                    self.transcription = "Audio transcription failed"
            
            # Clean up temporary audio file
            audio_path.unlink()
            
        except Exception as e:
            logger.error(f"Error in audio extraction/transcription: {e}")
            self.transcription = "Audio processing failed"
        
        return self.transcription
    
    def analyze_frame_content(self, frame):
        """Analyze frame content and generate description."""
        # Simple analysis based on color distribution and edges
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        edges = cv2.Canny(gray, 50, 150)
        edge_density = np.mean(edges) / 255
        
        # Color analysis
        colors = np.mean(frame, axis=(0, 1))
        dominant_color = "blue" if colors[0] > max(colors[1], colors[2]) else \
                        "green" if colors[1] > colors[2] else "red"
        
        # Generate description
        if edge_density > 0.1:
            complexity = "complex scene with many details"
        elif edge_density > 0.05:
            complexity = "moderate complexity"
        else:
            complexity = "simple scene"
        
        return f"Frame shows {complexity} with {dominant_color} tones"
    
    def create_powerpoint(self):
        """Create PowerPoint presentation from extracted frames and content."""
        logger.info("Creating PowerPoint presentation...")
        
        # Create presentation
        prs = Presentation()
        
        # Add title slide
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = f"Video Analysis: {self.video_path.name}"
        subtitle.text = f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        # Add summary slide with transcription
        if self.transcription:
            content_slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(content_slide_layout)
            title = slide.shapes.title
            content = slide.placeholders[1]
            
            title.text = "Video Transcription Summary"
            content.text = self.transcription[:500] + "..." if len(self.transcription) > 500 else self.transcription
        
        # Add frame slides
        for i, (frame, timestamp) in enumerate(zip(self.extracted_frames, self.frame_timestamps)):
            # Use content with caption layout
            slide_layout = prs.slide_layouts[5]  # Blank layout
            slide = prs.slides.add_slide(slide_layout)
            
            # Add title
            title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
            title_frame = title_shape.text_frame
            title_frame.text = f"Frame {i+1} - {timestamp:.2f}s"
            title_para = title_frame.paragraphs[0]
            title_para.font.size = Pt(self.config['title_font_size'])
            title_para.font.bold = True
            
            # Save frame as image
            frame_path = self.output_path.parent / f"frame_{i+1}.png"
            cv2.imwrite(str(frame_path), frame)
            
            # Add image to slide
            left = Inches(1)
            top = Inches(1.5)
            height = Inches(4)
            pic = slide.shapes.add_picture(str(frame_path), left, top, height=height)
            
            # Add frame analysis text
            left = Inches(1)
            top = Inches(6)
            width = Inches(8)
            height = Inches(1.5)
            
            text_box = slide.shapes.add_textbox(left, top, width, height)
            text_frame = text_box.text_frame
            text_frame.text = self.analyze_frame_content(frame)
            
            # Clean up temporary image file
            frame_path.unlink()
        
        # Save presentation
        prs.save(self.output_path)
        logger.info(f"PowerPoint presentation saved to: {self.output_path}")
    
    def process(self):
        """Run the complete video to PowerPoint conversion pipeline."""
        logger.info(f"Starting video to PowerPoint conversion for: {self.video_path}")
        
        # Extract key frames
        self.extract_key_frames()
        
        # Extract and transcribe audio
        self.extract_audio_and_transcribe()
        
        # Create PowerPoint presentation
        self.create_powerpoint()
        
        logger.info("Conversion completed successfully!")
        return {
            'output_path': str(self.output_path),
            'frames_extracted': len(self.extracted_frames),
            'transcription_length': len(self.transcription),
            'timestamps': self.frame_timestamps
        }

def main():
    """Main function to run the video to PowerPoint converter."""
    parser = argparse.ArgumentParser(description='Convert video to PowerPoint presentation')
    parser.add_argument('video_path', help='Path to input video file')
    parser.add_argument('output_path', help='Path for output PowerPoint file')
    parser.add_argument('--max-frames', type=int, default=10, help='Maximum number of frames to extract')
    parser.add_argument('--no-audio', action='store_true', help='Skip audio extraction and transcription')
    parser.add_argument('--config', help='Path to JSON configuration file')
    
    args = parser.parse_args()
    
    # Load configuration
    config = None
    if args.config:
        with open(args.config, 'r') as f:
            config = json.load(f)
    else:
        config = {
            'max_frames': args.max_frames,
            'audio_extraction': not args.no_audio,
            'frame_analysis': True,
            'slide_layout': 'image_with_text',
            'font_size': 16,
            'title_font_size': 24
        }
    
    # Create converter and process
    converter = VideoToPPTConverter(args.video_path, args.output_path, config)
    result = converter.process()
    
    print(f"Conversion completed successfully!")
    print(f"Output file: {result['output_path']}")
    print(f"Frames extracted: {result['frames_extracted']}")
    print(f"Transcription length: {result['transcription_length']} characters")

if __name__ == "__main__":
    main()
