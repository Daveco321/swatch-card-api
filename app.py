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

TARGET_SIZE = 100
COL_WIDTH_UNITS = 18

def process_image(url, index):
    if not url:
        return None
    try:
        response = requests.get(url, timeout=10)
        if response.status_code != 200:
            return None
        img_data = BytesIO(response.content)
        return {
            'index': index,
            'image_data': img_data,
            'x_scale': 0.5,
            'y_scale': 0.5,
            'x_offset': 5,
            'y_offset': 5,
            'url': url
        }
    except Exception as e:
        print(f"Error processing image {index}: {e}")
        return None

@app.route('/api/export-excel', methods=['POST'])
def export_excel():
    try:
        data = request.json
        swatches = data.get('swatches', [])
        card_info = data.get('cardInfo', {})
        
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
        
        worksheet.hide_gridlines(2)
        worksheet.freeze_panes(1, 0)
        
        headers = ['Image', 'Style #', 'Brand', 'Fit', 'Fabric Code', 'Fabrication', 'Color Name', 'Delivery', 'PO Ref']
        worksheet.set_row(0, 25)
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, fmt_header)
        
        worksheet.set_column(0, 0, COL_WIDTH_UNITS)
        worksheet.set_column(1, 1, 22)
        worksheet.set_column(2, 2, 18)
        worksheet.set_column(3, 3, 14)
        worksheet.set_column(4, 4, 12)
        worksheet.set_column(5, 5, 35)
        worksheet.set_column(6, 6, 18)
        worksheet.set_column(7, 7, 18)
        worksheet.set_column(8, 8, 12)
        
        worksheet.set_default_row(112.5)
        
        print(f"Processing {len(swatches)} swatches...")
        processed_images = {}
        
        with ThreadPoolExecutor(max_workers=10) as executor:
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
                except Exception:
                    pass
        
        print(f"Downloaded {len(processed_images)}/{len(swatches)} images")
        
        for row_num, swatch in enumerate(swatches):
            excel_row = row_num + 1
            is_even = (row_num % 2 == 1)
            
            fmt_cell = fmt_cell_even if is_even else fmt_cell_odd
            fmt_style = fmt_style_even if is_even else fmt_style_odd
            fmt_po = fmt_po_even if is_even else fmt_po_odd
            
            worksheet.write(excel_row, 0, '', fmt_cell)
            worksheet.write(excel_row, 1, swatch.get('styleNumber', ''), fmt_style)
            worksheet.write(excel_row, 2, swatch.get('brand', ''), fmt_cell)
            worksheet.write(excel_row, 3, swatch.get('fit', ''), fmt_cell)
            worksheet.write(excel_row, 4, swatch.get('fabricCode', ''), fmt_cell)
            worksheet.write(excel_row, 5, swatch.get('fabrication', ''), fmt_cell)
            worksheet.write(excel_row, 6, swatch.get('colorName', ''), fmt_cell)
            worksheet.write(excel_row, 7, swatch.get('delivery', ''), fmt_cell)
            worksheet.write(excel_row, 8, swatch.get('poRef', ''), fmt_po)
            
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
                    print(f"Error inserting image: {e}")
                    worksheet.write(excel_row, 0, "Error", fmt_cell)
            else:
                worksheet.write(excel_row, 0, "No Image", fmt_cell)
        
        footer_row = len(swatches) + 2
        fmt_footer = workbook.add_format({'font_size': 9, 'font_color': '#666666'})
        worksheet.write(footer_row, 1, '320 West 37th Street, 3rd floor, New York, NY 10018 | Tel 212-697-1660', fmt_footer)
        
        workbook.close()
        
        output.seek(0)
        filename = f"SwatchCard_{card_info.get('poRef', 'Export')}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        print(f"Error generating Excel: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'swatch-card-api'})

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        'service': 'Swatch Card API',
        'endpoints': {
            '/api/export-excel': 'POST - Generate Excel file',
            '/api/health': 'GET - Health check'
        }
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
