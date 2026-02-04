from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import xlsxwriter
import requests
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor, as_completed
import os

app = Flask(__name__)
CORS(app)

STYLE_CONFIG = {
    'header_bg': '#1E3A5F',
    'header_text': '#FFFFFF',
    'row_bg_odd': '#FFFFFF',
    'row_bg_even': '#F0F4F8',
    'border_color': '#E2E8F0',
    'po_ref_color': '#DC2626',
}

# Image sizing - 150x150 pixels
IMAGE_SIZE = 150
ROW_HEIGHT = 112.5  # Points (150px * 0.75)
COL_WIDTH = 21      # Characters (~150px)

def process_image(url, index):
    if not url:
        print(f"Image {index}: No URL provided")
        return None
    try:
        clean_url = url.replace(' ', '%20').replace('+', '%20')
        print(f"Image {index}: Fetching {clean_url[:80]}...")
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(clean_url, timeout=15, headers=headers)
        
        if response.status_code != 200:
            print(f"Image {index}: Failed with status {response.status_code}")
            return None
        
        content_type = response.headers.get('content-type', '')
        if 'image' not in content_type and len(response.content) < 1000:
            print(f"Image {index}: Not an image (content-type: {content_type})")
            return None
            
        img_data = BytesIO(response.content)
        print(f"Image {index}: Success ({len(response.content)} bytes)")
        
        return {
            'index': index,
            'image_data': img_data,
            'url': url
        }
    except Exception as e:
        print(f"Image {index}: Error - {str(e)}")
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
            'font_color': STYLE_CONFIG['header_text'],
            'bg_color': STYLE_CONFIG['header_bg'],
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'font_size': 11,
        })
        
        base_props = {
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True,
            'font_size': 10,
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
        worksheet.set_column(0, 0, COL_WIDTH)      # Image - 150px wide
        worksheet.set_column(1, 1, 18)             # Style #
        worksheet.set_column(2, 2, 15)             # Brand
        worksheet.set_column(3, 3, 12)             # Fit
        worksheet.set_column(4, 4, 12)             # Fabric Code
        worksheet.set_column(5, 5, 32)             # Fabrication
        worksheet.set_column(6, 6, 16)             # Color Name
        worksheet.set_column(7, 7, 16)             # Delivery
        worksheet.set_column(8, 8, 12)             # PO Ref
        
        # Set default row height for image rows
        worksheet.set_default_row(ROW_HEIGHT)
        
        # Download images in parallel
        print(f"Downloading images...")
        processed_images = {}
        
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {
                executor.submit(process_image, s.get('imageUrl'), idx): idx 
                for idx, s in enumerate(swatches)
            }
            for future in as_completed(futures):
                idx = futures[future]
                try:
                    result = future.result()
                    if result:
                        processed_images[idx] = result
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
            
            # Insert image - sized to 150x150
            img_data = processed_images.get(row_num)
            if img_data:
                try:
                    worksheet.insert_image(excel_row, 0, "img.png", {
                        'image_data': img_data['image_data'],
                        'x_offset': 2,
                        'y_offset': 2,
                        'x_scale': 0.25,
                        'y_scale': 0.25,
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
    return jsonify({'status': 'ok', 'service': 'swatch-card-api', 'version': '1.2'})

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'service': 'Swatch Card API',
        'version': '1.2',
        'endpoints': {
            '/api/export-excel': 'POST - Generate Excel file',
            '/api/health': 'GET - Health check'
        }
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
