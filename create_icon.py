#!/usr/bin/env python3
"""
Create a professional data quality icon for the DQA application.
"""

from PIL import Image, ImageDraw, ImageFont
import numpy as np
import os

def create_data_quality_icon():
    """Create a professional data quality icon with multiple sizes."""
    
    # Create different sizes for the icon
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = []
    
    for size in sizes:
        # Create a new image with transparent background
        img = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        width, height = size
        
        # Professional color scheme: Blue (trust) + Green (quality)
        primary_color = (25, 118, 210)  # Material Blue 600
        secondary_color = (76, 175, 80)  # Material Green 600
        accent_color = (255, 193, 7)     # Material Amber 500
        white = (255, 255, 255, 255)
        
        # Draw background circle/shield
        if width >= 64:
            # Draw shield shape for larger icons
            shield_points = [
                (width * 0.5, height * 0.1),   # Top center
                (width * 0.8, height * 0.3),   # Top right
                (width * 0.8, height * 0.7),   # Bottom right
                (width * 0.5, height * 0.9),   # Bottom center
                (width * 0.2, height * 0.7),   # Bottom left
                (width * 0.2, height * 0.3),   # Top left
            ]
            draw.polygon(shield_points, fill=primary_color, outline=white, width=max(1, width//64))
        else:
            # Simple circle for small icons
            margin = width // 8
            draw.ellipse([margin, margin, width-margin, height-margin], 
                        fill=primary_color, outline=white, width=1)
        
        # Draw data visualization elements (chart bars)
        if width >= 32:
            # Draw chart bars
            bar_width = width // 6
            bar_spacing = bar_width // 2
            base_y = height * 0.7
            
            # Bar heights (representing data quality metrics)
            bar_heights = [height * 0.5, height * 0.7, height * 0.4, height * 0.6]
            bar_colors = [secondary_color, accent_color, secondary_color, accent_color]
            
            start_x = (width - (4 * bar_width + 3 * bar_spacing)) // 2
            
            for i in range(4):
                x1 = start_x + i * (bar_width + bar_spacing)
                y1 = base_y - bar_heights[i]
                x2 = x1 + bar_width
                y2 = base_y
                
                draw.rectangle([x1, y1, x2, y2], fill=bar_colors[i])
        
        # Draw checkmark (quality symbol)
        if width >= 48:
            check_size = width // 4
            center_x = width // 2
            center_y = height // 2 if width < 64 else height * 0.4
            
            # Draw checkmark
            check_points = [
                (center_x - check_size//2, center_y),
                (center_x - check_size//4, center_y + check_size//3),
                (center_x + check_size//2, center_y - check_size//3),
            ]
            draw.line(check_points, fill=white, width=max(2, width//32), joint="curve")
        
        # Add "DQA" text for larger icons
        if width >= 128:
            try:
                # Try to use a professional font
                font_size = width // 8
                font = ImageFont.truetype("arial.ttf", font_size)
                text = "DQA"
                text_bbox = draw.textbbox((0, 0), text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]
                
                text_x = (width - text_width) // 2
                text_y = height * 0.8 if width >= 256 else height * 0.75
                
                draw.text((text_x, text_y), text, font=font, fill=white)
            except:
                # Fallback if font not available
                pass
        
        images.append(img)
    
    return images

def save_as_ico(images, filename="icon.ico"):
    """Save images as ICO file."""
    # Convert all images to RGB mode for ICO format
    rgb_images = []
    for img in images:
        # Create a copy in RGB mode with white background
        rgb_img = Image.new('RGB', img.size, (255, 255, 255))
        rgb_img.paste(img, mask=img.split()[3] if img.mode == 'RGBA' else None)
        rgb_images.append(rgb_img)
    
    # Save as ICO
    rgb_images[0].save(
        filename,
        format='ICO',
        sizes=[img.size for img in rgb_images],
        append_images=rgb_images[1:]
    )
    
    print(f"✓ Professional icon saved as {filename}")
    print(f"  Sizes: {[img.size for img in rgb_images]}")

def main():
    """Main function to create the icon."""
    print("Creating professional data quality icon...")
    
    try:
        # Create the icon
        images = create_data_quality_icon()
        
        # Save as ICO
        save_as_ico(images, "icon.ico")
        
        # Also save a preview
        images[0].save("icon_preview.png", "PNG")
        print("✓ Preview saved as icon_preview.png")
        
        print("\n🎨 Icon Design:")
        print("   • Shield shape representing protection/security")
        print("   • Chart bars representing data analytics")
        print("   • Checkmark representing quality assurance")
        print("   • Blue/Green color scheme (professional/healthcare)")
        print("   • DQA text on larger versions")
        
    except Exception as e:
        print(f"Error creating icon: {e}")
        print("\nUsing fallback method...")
        
        # Create a simple fallback icon
        create_simple_icon()

def create_simple_icon():
    """Create a simple fallback icon."""
    from PIL import Image, ImageDraw
    
    # Create a simple 256x256 icon
    img = Image.new('RGB', (256, 256), (25, 118, 210))
    draw = ImageDraw.Draw(img)
    
    # Draw a simple checkmark
    draw.line([(80, 150), (110, 180), (180, 100)], fill=(76, 175, 80), width=20)
    
    # Draw DQA text
    try:
        from PIL import ImageFont
        font = ImageFont.truetype("arial.ttf", 40)
        draw.text((85, 50), "DQA", fill=(255, 255, 255), font=font)
    except:
        pass
    
    img.save("icon.ico", format='ICO')
    print("✓ Simple icon created as fallback")

if __name__ == "__main__":
    main()