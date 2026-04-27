#!/usr/bin/env python3
"""
GPT-5.4 Enhanced PPT Desensitization Tool
Main application for intelligent PowerPoint data redaction
"""

import argparse
import sys
import os
from pathlib import Path
import json
import logging
from datetime import datetime

# Import our modules
from ppt_desensitizer import GPTEnhancedDesensitizer
from ppt_processor import PowerPointProcessor, PPTXMLProcessor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ppt_desensitization.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class PPTDesensitizationApp:
    """
    Main application for PPT desensitization with GPT-5.4 enhancement
    """
    
    def __init__(self, config_path: str = None):
        self.config_path = config_path or 'desensitization_config.json'
        self.desensitizer = GPTEnhancedDesensitizer(self.config_path)
        self.processor = PowerPointProcessor(self.desensitizer)
        self.xml_processor = PPTXMLProcessor(self.desensitizer)
        self.start_time = None
        self.end_time = None
    
    def process_file(self, input_path: str, output_path: str = None) -> bool:
        """
        Process a single PPT file
        """
        try:
            self.start_time = datetime.now()
            
            # Validate input file
            if not os.path.exists(input_path):
                logger.error(f"Input file not found: {input_path}")
                return False
            
            if not input_path.lower().endswith(('.pptx', '.ppt')):
                logger.error(f"Unsupported file format: {input_path}")
                return False
            
            # Generate output path if not provided
            if not output_path:
                input_file = Path(input_path)
                output_path = str(input_file.parent / f"{input_file.stem}_desensitized{input_file.suffix}")
            
            logger.info(f"Starting desensitization of: {input_path}")
            logger.info(f"Output will be saved to: {output_path}")
            
            # Process the presentation
            success = self.processor.process_presentation(input_path, output_path)
            
            if success:
                self.end_time = datetime.now()
                processing_time = (self.end_time - self.start_time).total_seconds()
                
                # Generate summary report
                summary = self._generate_summary_report(input_path, output_path, processing_time)
                
                logger.info("Desensitization completed successfully!")
                logger.info(f"Processing time: {processing_time:.2f} seconds")
                logger.info(f"Summary: {json.dumps(summary, indent=2, ensure_ascii=False)}")
                
                return True
            else:
                logger.error("Desensitization failed!")
                return False
                
        except Exception as e:
            logger.error(f"Unexpected error during processing: {e}")
            return False
    
    def process_directory(self, input_dir: str, output_dir: str = None) -> Dict:
        """
        Process all PPT files in a directory
        """
        results = {
            'total_files': 0,
            'successful': 0,
            'failed': 0,
            'skipped': 0,
            'processing_times': []
        }
        
        try:
            input_path = Path(input_dir)
            if not input_path.exists():
                logger.error(f"Input directory not found: {input_dir}")
                return results
            
            # Create output directory if specified
            if output_dir:
                output_path = Path(output_dir)
                output_path.mkdir(parents=True, exist_ok=True)
            else:
                output_path = input_path / "desensitized"
                output_path.mkdir(exist_ok=True)
            
            # Find all PPT files
            ppt_files = list(input_path.glob("*.pptx")) + list(input_path.glob("*.ppt"))
            results['total_files'] = len(ppt_files)
            
            logger.info(f"Found {len(ppt_files)} PPT files to process")
            
            for i, ppt_file in enumerate(ppt_files, 1):
                logger.info(f"Processing file {i}/{len(ppt_files)}: {ppt_file.name}")
                
                output_file = output_path / f"{ppt_file.stem}_desensitized{ppt_file.suffix}"
                
                file_start = datetime.now()
                success = self.process_file(str(ppt_file), str(output_file))
                file_end = datetime.now()
                
                processing_time = (file_end - file_start).total_seconds()
                results['processing_times'].append(processing_time)
                
                if success:
                    results['successful'] += 1
                    logger.info(f"✓ Successfully processed: {ppt_file.name}")
                else:
                    results['failed'] += 1
                    logger.error(f"✗ Failed to process: {ppt_file.name}")
            
            # Calculate average processing time
            if results['processing_times']:
                results['average_processing_time'] = sum(results['processing_times']) / len(results['processing_times'])
            
            logger.info(f"Batch processing completed!")
            logger.info(f"Total files: {results['total_files']}")
            logger.info(f"Successful: {results['successful']}")
            logger.info(f"Failed: {results['failed']}")
            
            return results
            
        except Exception as e:
            logger.error(f"Error during batch processing: {e}")
            return results
    
    def _generate_summary_report(self, input_path: str, output_path: str, processing_time: float) -> Dict:
        """Generate processing summary report"""
        processor_summary = self.processor.get_processing_summary()
        desensitizer_stats = self.desensitizer.get_statistics()
        
        summary = {
            'input_file': input_path,
            'output_file': output_path,
            'processing_time_seconds': processing_time,
            'timestamp': datetime.now().isoformat(),
            'processor_summary': processor_summary,
            'desensitizer_stats': desensitizer_stats,
            'tool_version': '2.0.0',
            'ai_model': 'GPT-5.4 Enhanced',
            'redaction_accuracy': processor_summary.get('redaction_rate', 0) * 100,
            'preservation_accuracy': processor_summary.get('preservation_rate', 0) * 100
        }
        
        # Save summary report
        summary_path = output_path.replace('.pptx', '_summary.json').replace('.ppt', '_summary.json')
        try:
            with open(summary_path, 'w', encoding='utf-8') as f:
                json.dump(summary, f, ensure_ascii=False, indent=2)
            logger.info(f"Summary report saved: {summary_path}")
        except Exception as e:
            logger.warning(f"Failed to save summary report: {e}")
        
        return summary
    
    def validate_configuration(self) -> bool:
        """Validate the current configuration"""
        try:
            config_valid = self.desensitizer.get_statistics()
            logger.info("Configuration validation passed")
            logger.info(f"Available rules: {config_valid['rules_applied']}")
            logger.info(f"AI enhancement: {'Enabled' if config_valid['ai_enhancement_enabled'] else 'Disabled'}")
            return True
        except Exception as e:
            logger.error(f"Configuration validation failed: {e}")
            return False

def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='GPT-5.4 Enhanced PPT Desensitization Tool',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process single file
  python ppt_desensitization_app.py input.pptx
  
  # Process single file with custom output
  python ppt_desensitization_app.py input.pptx -o output.pptx
  
  # Process directory
  python ppt_desensitization_app.py /path/to/ppt/files/
  
  # Process directory with custom config
  python ppt_desensitization_app.py /path/to/ppt/files/ -c custom_config.json
  
  # Batch processing
  python ppt_desensitization_app.py /path/to/input/ -d /path/to/output/
        """
    )
    
    parser.add_argument('input', help='Input PPT file or directory')
    parser.add_argument('-o', '--output', help='Output file path (for single file processing)')
    parser.add_argument('-d', '--output-dir', help='Output directory (for batch processing)')
    parser.add_argument('-c', '--config', help='Configuration file path', default='desensitization_config.json')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose logging')
    parser.add_argument('--validate', action='store_true', help='Validate configuration only')
    
    args = parser.parse_args()
    
    # Set logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Initialize application
    app = PPTDesensitizationApp(args.config)
    
    # Validate configuration if requested
    if args.validate:
        success = app.validate_configuration()
        sys.exit(0 if success else 1)
    
    # Check if input is file or directory
    input_path = Path(args.input)
    
    if input_path.is_file():
        # Process single file
        success = app.process_file(args.input, args.output)
        sys.exit(0 if success else 1)
    
    elif input_path.is_dir():
        # Process directory
        results = app.process_directory(args.input, args.output_dir)
        
        # Print final summary
        print("\n" + "="*50)
        print("PROCESSING SUMMARY")
        print("="*50)
        print(f"Total files processed: {results['total_files']}")
        print(f"Successful: {results['successful']}")
        print(f"Failed: {results['failed']}")
        print(f"Skipped: {results['skipped']}")
        
        if results['processing_times']:
            avg_time = sum(results['processing_times']) / len(results['processing_times'])
            print(f"Average processing time: {avg_time:.2f} seconds")
        
        sys.exit(0 if results['failed'] == 0 else 1)
    
    else:
        logger.error(f"Input path does not exist: {args.input}")
        sys.exit(1)

if __name__ == "__main__":
    main()