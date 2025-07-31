import pyperclip
import subprocess
import tempfile
import os
import re
import sys
from datetime import datetime

def convert_math_blocks(md_content):
    """
    Convert ```math blocks to appropriate LaTeX format that pandoc can recognize.
    
    For single line math expressions, convert to $...$ format.
    For multi-line math expressions, convert to $$...$$ format.
    """
    def replace_math_block(match):
        math_content = match.group(1).strip()
        
        # Check if it's a single line or multi-line
        lines = math_content.split('\n')
        
        # Remove any empty lines at the beginning or end
        lines = [line for line in lines if line.strip()]
        
        if len(lines) == 1:
            # Single line - use $...$ format
            return f"${lines[0]}$"
        else:
            # Multi-line - use $$...$$ format
            inner_content = '\n'.join(lines)
            return f"$${inner_content}$$"
    
    # Pattern to match ```math blocks
    pattern = r'```math\s*(.*?)\s*```'
    
    # Replace all matches
    converted_content = re.sub(pattern, replace_math_block, md_content, flags=re.DOTALL)
    
    return converted_content

def is_markdown(content):
    """
    Check if content is Markdown by looking for common Markdown patterns.
    This is a more robust implementation.
    """
    # If content is empty, it's not Markdown
    if not content.strip():
        return False
    
    # Common Markdown patterns
    patterns = [
        r'^#{1,6}\s',  # Headers
        r'\*\*.*?\*\*',  # Bold with **
        r'\*.*?\*',  # Italic with *
        r'`[^`]*`',  # Inline code
        r'^\s*[-+*]\s',  # Unordered lists
        r'^\s*\d+\.\s',  # Ordered lists
        r'\[.*?\]\(.*?\)',  # Links
        r'!\[.*?\]\(.*?\)',  # Images
        r'^\s*>',  # Blockquotes
        r'```',  # Code blocks
        r'^\s*[-*_]{3,}\s*$',  # Horizontal rules
    ]
    
    lines = content.split('\n')
    matches = 0
    
    # Check each line for Markdown patterns
    for line in lines:
        for pattern in patterns:
            if re.search(pattern, line):
                matches += 1
                # If we find several Markdown patterns, it's likely Markdown
                if matches >= 2:
                    return True
    
    # Additional check for typical Markdown file characteristics
    # Check if there are headers or list patterns
    if re.search(r'^#{1,6}\s', content, re.MULTILINE) or \
       re.search(r'^\s*[-+*]\s', content, re.MULTILINE) or \
       re.search(r'^\s*\d+\.\s', content, re.MULTILINE):
        return True
    
    return False

def convert_clipboard_to_docx():
    """
    Convert clipboard content to DOCX file with timestamped filename.
    """
    try:
        # Read content from clipboard
        print("Reading content from clipboard...")
        clipboard_content = pyperclip.paste()
        
        # Check if content is empty
        if not clipboard_content.strip():
            print("Error: Clipboard is empty")
            return False
        
        print(f"Clipboard content length: {len(clipboard_content)} characters")
        
        # Check if content is Markdown
        print("Checking if content is Markdown...")
        if not is_markdown(clipboard_content):
            print("Error: Content is not in Markdown format")
            return False
        
        print("Content identified as Markdown")
        
        # Convert LaTeX formulas in ```math blocks to formats pandoc can recognize
        print("Converting LaTeX formulas...")
        converted_content = convert_math_blocks(clipboard_content)
        
        # Create a timestamped filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"markdown_conversion_{timestamp}.docx"
        
        # Get the directory where this script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(script_dir, output_filename)
        
        print(f"Creating DOCX file at: {output_path}")
        
        # Create temporary file for the converted markdown content
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as md_file:
            md_file.write(converted_content)
            md_path = md_file.name
        
        print(f"Created temporary Markdown file: {md_path}")
        
        try:
            # Run pandoc to convert MD to DOCX
            print("Running pandoc conversion...")
            result = subprocess.run([
                'pandoc', 
                md_path, 
                '-o', 
                output_path
            ], capture_output=True, text=True, timeout=30)
            
            if result.returncode != 0:
                raise Exception(f"Pandoc conversion failed: {result.stderr}")
            
            # Check if the DOCX file was created
            if not os.path.exists(output_path):
                raise Exception("DOCX file was not created")
                
            print("Conversion completed successfully!")
            print(f"Output file: {output_path}")
            
            # Put the file path on the clipboard
            pyperclip.copy(output_path)
            print("File path copied to clipboard")
            
            return True
            
        except subprocess.TimeoutExpired:
            raise Exception("Pandoc conversion timed out")
        except Exception:
            # Clean up files if conversion failed
            if os.path.exists(output_path):
                os.unlink(output_path)
            raise
        finally:
            # Clean up temporary MD file
            if os.path.exists(md_path):
                os.unlink(md_path)
                print(f"Cleaned up temporary file: {md_path}")
    
    except Exception as e:
        print(f"Error: {str(e)}")
        return False

def main():
    print("Markdown to Word Converter (Clipboard Version)")
    print("=" * 50)
    
    success = convert_clipboard_to_docx()
    
    if success:
        print("\nConversion completed successfully!")
        input("Press Enter to exit...")
        return 0
    else:
        print("\nConversion failed!")
        input("Press Enter to exit...")
        return 1

if __name__ == "__main__":
    sys.exit(main())
