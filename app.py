from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import xlsxwriter
import requests
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
from PIL import Image as PilImage
from PIL import ImageOps
import os

app = Flask(__name__)
CORS(app)

# Image sizing - matching employee's exact settings
TARGET_W = 150
TARGET_H = 150
COL_WIDTH_UNITS = 22
ROW_HEIGHT = 112.5

STYLE_CONFIG = {
    'header_bg': '#1E3A5F',
    'header_text': '#FFFFFF',
    'row_bg_odd': '#FFFFFF',
    'row_bg_even': '#F0F4F8',
    'border_color': '#E2E8F0',
    'po_ref_color': '#DC2626',
    'font_name': 'Calibri'
}

def process_single_image(url, target_width, target_height):
    """EMPLOYEE'S EXACT IMAGE PROCESSING CODE"""
    if not (isinstance(url, str) and url.startswith('http')):
        return None
    
    # Clean URL
    clean_url = url.replace(' ', '%20').replace('+', '%20')
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    
    try:
        response = requests.get(clean_url, headers=headers, timeout=15)
        
        content_type = response.headers.get('Content-Type', '').lower()
        if 'image' not in content_type:
            print(f"Not an image: {content_type}")
            return None
        if response.status_code != 200:
            print(f"Failed with status {response.status_code}")
            return None
            
        image_data = BytesIO(response.content)
        
        with PilImage.open(image_data) as im:
            im = ImageOps.exif_transpose(im)
            
            im.thumbnail((target_width * 2, target_height * 2), PilImage.Resampling.LANCZOS)
            
            output_format = "PNG"
            if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                output_format = "PNG"
            else:
                if im.mode != "RGB":
                    im = im.convert("RGB")
                output_format = "JPEG"
            
            processed_image_data = BytesIO()
            im.save(processed_image_data, format=output_format, quality=85, optimize=True)
            processed_image_data.seek(0)
            
            orig_w, orig_h = im.size
        
        width_ratio = target_width / orig_w
        height_ratio = target_height / orig_h
        scale_factor = min(width_ratio, height_ratio)
        
        final_w = orig_w * scale_factor
        final_h = orig_h * scale_factor
        
        x_offset = (target_width - final_w) / 2
        y_offset = (target_height - final_h) / 2
        
        return {
            'image_data': processed_image_data,
            'x_scale': scale_factor,
            'y_scale': scale_factor,
            'x_offset': x_offset,
            'y_offset': y_offset,
            'object_position': 1,
            'url': url
        }
    except Exception as e:
        print(f"Image error: {str(e)}")
        return None

@app.route('/api/export-excel', methods=['POST'])
def export_excel():
    try:
        data = request.json
        swatches = data.get('swatches', [])
        card_info = data.get('cardInfo', {})
        
        print(f"=== Excel Export Request ===")
        print(f"Swatches count: {len(swatches)}")
        
        if not swatches:
            return jsonify({'error': 'No swatches provided'}), 400
        
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        
        workbook.set_properties({
            'title': f"Swatch Card - {card_info.get('poRef', 'Export')}",
            'author': 'Swatch Card Builder',
            'company': 'HalfPrice',
        })
        
        worksheet = workbook.add_worksheet('Swatch Card')
        
        # Formats
        fmt_header = workbook.add_format({
            'bold': True,
            'font_name': STYLE_CONFIG['font_name'],
            'font_size': 11,
            'font_color': STYLE_CONFIG['header_text'],
            'bg_color': STYLE_CONFIG['header_bg'],
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'border_color': STYLE_CONFIG['border_color'],
        })
        
        base_props = {
            'font_name': STYLE_CONFIG['font_name'],
            'font_size': 10,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'border_color': STYLE_CONFIG['border_color'],
            'text_wrap': True,
        }
        
        fmt_cell_odd = workbook.add_format({**base_props, 'bg_color': STYLE_CONFIG['row_bg_odd']})
        fmt_cell_even = workbook.add_format({**base_props, 'bg_color': STYLE_CONFIG['row_bg_even']})
        
        fmt_style_odd = workbook.add_format({
            **base_props, 
            'bg_color': STYLE_CONFIG['row_bg_odd'],
            'bold': True,
            'font_name': 'Consolas',
            'font_size': 11,
        })
        fmt_style_even = workbook.add_format({
            **base_props, 
            'bg_color': STYLE_CONFIG['row_bg_even'],
            'bold': True,
            'font_name': 'Consolas',
            'font_size': 11,
        })
        
        fmt_po_odd = workbook.add_format({
            **base_props, 
            'bg_color': STYLE_CONFIG['row_bg_odd'],
            'bold': True,
            'font_color': STYLE_CONFIG['po_ref_color'],
        })
        fmt_po_even = workbook.add_format({
            **base_props, 
            'bg_color': STYLE_CONFIG['row_bg_even'],
            'bold': True,
            'font_color': STYLE_CONFIG['po_ref_color'],
        })
        
        # Hide gridlines and freeze header
        worksheet.hide_gridlines(2)
        worksheet.freeze_panes(1, 0)
        
        # Headers
        headers = ['Image', 'Style #', 'Brand', 'Fit', 'Fabric Code', 'Fabrication', 'Color Name', 'Delivery', 'PO Ref']
        worksheet.set_row(0, 25)
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, fmt_header)
        
        # Column widths
        worksheet.set_column(0, 0, COL_WIDTH_UNITS)  # Image - 150px wide
        worksheet.set_column(1, 1, 18)               # Style #
        worksheet.set_column(2, 2, 15)               # Brand
        worksheet.set_column(3, 3, 12)               # Fit
        worksheet.set_column(4, 4, 12)               # Fabric Code
        worksheet.set_column(5, 5, 32)               # Fabrication
        worksheet.set_column(6, 6, 16)               # Color Name
        worksheet.set_column(7, 7, 16)               # Delivery
        worksheet.set_column(8, 8, 12)               # PO Ref
        
        # Set default row height for image rows
        worksheet.set_default_row(ROW_HEIGHT)
        
        # Download images in parallel
        print(f"Downloading images...")
        processed_images = {}
        
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {
                executor.submit(process_single_image, s.get('imageUrl'), TARGET_W, TARGET_H): idx 
                for idx, s in enumerate(swatches)
            }
            for future in as_completed(futures):
                idx = futures[future]
                try:
                    result = future.result()
                    if result:
                        processed_images[idx] = result
                        print(f"✓ Image {idx+1}/{len(swatches)}")
                except Exception as e:
                    print(f"Future error for {idx}: {e}")
        
        print(f"Downloaded {len(processed_images)}/{len(swatches)} images")
        
        # Write data rows
        for row_num, swatch in enumerate(swatches):
            excel_row = row_num + 1
            is_even = (row_num % 2 == 1)
            
            fmt_cell = fmt_cell_even if is_even else fmt_cell_odd
            fmt_style = fmt_style_even if is_even else fmt_style_odd
            fmt_po = fmt_po_even if is_even else fmt_po_odd
            
            # Write cell data
            worksheet.write(excel_row, 0, '', fmt_cell)  # Image placeholder
            worksheet.write(excel_row, 1, swatch.get('styleNumber', ''), fmt_style)
            worksheet.write(excel_row, 2, swatch.get('brand', ''), fmt_cell)
            worksheet.write(excel_row, 3, swatch.get('fit', ''), fmt_cell)
            worksheet.write(excel_row, 4, swatch.get('fabricCode', ''), fmt_cell)
            worksheet.write(excel_row, 5, swatch.get('fabrication', ''), fmt_cell)
            worksheet.write(excel_row, 6, swatch.get('colorName', ''), fmt_cell)
            worksheet.write(excel_row, 7, swatch.get('delivery', ''), fmt_cell)
            worksheet.write(excel_row, 8, swatch.get('poRef', ''), fmt_po)
            
            # Insert image - using employee's exact method
            img_data = processed_images.get(row_num)
            if img_data:
                try:
                    worksheet.insert_image(excel_row, 0, "img.png", {
                        'image_data': img_data['image_data'],
                        'x_scale': img_data['x_scale'],
                        'y_scale': img_data['y_scale'],
                        'x_offset': img_data['x_offset'],
                        'y_offset': img_data['y_offset'],
                        'object_position': 1,
                    })
                except Exception as e:
                    print(f"Insert image error row {row_num}: {e}")
                    worksheet.write(excel_row, 0, "Error", fmt_cell)
            else:
                worksheet.write(excel_row, 0, "No Image", fmt_cell)
        
        # Footer
        footer_row = len(swatches) + 2
        fmt_footer = workbook.add_format({'font_size': 9, 'font_color': '#666666'})
        worksheet.write(footer_row, 1, '320 West 37th Street, 3rd floor, New York, NY 10018 | Tel 212-697-1660', fmt_footer)
        
        workbook.close()
        
        output.seek(0)
        filename = f"SwatchCard_{card_info.get('poRef', 'Export')}.xlsx"
        
        print(f"=== Export complete: {filename} ===")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        print(f"Export error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'swatch-card-api', 'version': '2.0'})

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'service': 'Swatch Card API',
        'version': '2.0',
        'endpoints': {
            '/api/export-excel': 'POST - Generate Excel file',
            '/api/health': 'GET - Health check'
        }
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
