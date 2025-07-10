#!/usr/bin/env python3
"""
Create proper PNG icons for Office Add-in
"""

from PIL import Image, ImageDraw, ImageFont
import os

def create_icon(size, filename):
    """Create a simple icon with specified size."""
    # Create image with transparent background
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Draw a brain emoji-like icon
    # Background circle
    margin = size // 8
    draw.ellipse([margin, margin, size-margin, size-margin], 
                 fill=(66, 133, 244, 255), outline=(55, 110, 200, 255), width=2)
    
    # Brain icon (simplified)
    center = size // 2
    brain_size = size // 3
    
    # Draw brain-like shape
    draw.ellipse([center - brain_size//2, center - brain_size//2, 
                  center + brain_size//2, center + brain_size//2], 
                 fill=(255, 255, 255, 255))
    
    # Add some brain-like lines
    if size >= 32:
        line_width = max(1, size // 32)
        draw.arc([center - brain_size//3, center - brain_size//3,
                  center + brain_size//3, center + brain_size//3],
                 start=0, end=180, fill=(66, 133, 244, 255), width=line_width)
        draw.arc([center - brain_size//4, center - brain_size//4,
                  center + brain_size//4, center + brain_size//4],
                 start=180, end=360, fill=(66, 133, 244, 255), width=line_width)
    
    # Save the image
    img.save(f'assets/{filename}', 'PNG')
    print(f"Created {filename} ({size}x{size})")

def main():
    """Create all required icon sizes."""
    os.makedirs('assets', exist_ok=True)
    
    # Create icons in required sizes
    icon_sizes = [
        (16, 'icon-16.png'),
        (32, 'icon-32.png'),
        (64, 'icon-64.png'),
        (80, 'icon-80.png')
    ]
    
    for size, filename in icon_sizes:
        create_icon(size, filename)
    
    print("All icons created successfully!")

if __name__ == "__main__":
    main() 