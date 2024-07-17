from flask import Flask, request, jsonify, send_file
from oletools.olevba import VBA_Parser
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib import colors
import tempfile
import os
import re

app = Flask(__name__)

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and file.filename.endswith('.xlsm'):
        try:
            temp_file = tempfile.NamedTemporaryFile(delete=False)
            file.save(temp_file.name)

            vba_code = extract_vba(temp_file.name)

            if not vba_code:
                return jsonify({'error': 'No VBA code found or an error occurred'}), 400

            analysis = analyze_vba(vba_code)

            pdf_filename = generate_pdf_documentation(analysis)

            return jsonify({'pdf_url': f'/download/{pdf_filename}'})

        except Exception as e:
            return jsonify({'error': str(e)}), 500

    return jsonify({'error': 'Invalid file type, only .xlsm files are supported'}), 400

def extract_vba(file_path):
    try:
        vba_parser = VBA_Parser(file_path)
        vba_modules = vba_parser.extract_all_macros()

        if not vba_modules:
            return ""

        vba_code = ""
        for (_, _, vba_filename, vba_code_all) in vba_modules:
            vba_code += f"\n\n--- Module: {vba_filename} ---\n"
            vba_code += vba_code_all

        return vba_code
    except Exception as e:
        raise RuntimeError(f"An error occurred while extracting VBA code: {str(e)}")

def analyze_vba(vba_code):
    analysis = {}
    modules = re.split(r'--- Module: .+ ---', vba_code)
    module_names = re.findall(r'--- Module: (.+) ---', vba_code)

    for i, module in enumerate(modules[1:], 0):
        module_name = module_names[i]
        functions = re.findall(r'(?:Public |Private )?(?:Sub|Function)\s+(\w+)\s*\((.*?)\)', module, re.DOTALL)
        variables = re.findall(r'Dim\s+(\w+)\s+As\s+(\w+)', module)
        comments = re.findall(r"'(.+)$", module, re.MULTILINE)
        logic = extract_logic(module)

        analysis[module_name] = {
            'functions': functions,
            'variables': variables,
            'comments': comments,
            'logic': logic,
            'full_code': module.strip()
        }

    return analysis

def extract_logic(module_code):
    logic = []
    loops = re.findall(r'(For\s+.+\s+To\s+.+|Do\s+While\s+.+|Do\s+Until\s+.+|Loop\s+Until\s+.+)', module_code, re.IGNORECASE)
    logic.extend(loops)
    conditions = re.findall(r'If\s+.+\s+Then\s+.+', module_code, re.IGNORECASE)
    logic.extend(conditions)
    function_calls = re.findall(r'\bCall\s+(\w+)', module_code, re.IGNORECASE)
    logic.extend(function_calls)
    return logic

def generate_pdf_documentation(analysis):
    try:
        pdf_filename = 'vba_documentation.pdf'
        pdf_path = os.path.join(os.path.dirname(__file__), pdf_filename)
        
        doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                                rightMargin=72, leftMargin=72,
                                topMargin=72, bottomMargin=18)
        Story = []

        title_style = ParagraphStyle(name='Title', fontSize=22, leading=30, alignment=TA_JUSTIFY)
        heading1_style = ParagraphStyle(name='Heading1', fontSize=18, leading=20, alignment=TA_JUSTIFY)
        heading2_style = ParagraphStyle(name='Heading2', fontSize=14, leading=18, alignment=TA_JUSTIFY)
        normal_style = ParagraphStyle(name='Normal', fontSize=12, leading=14, alignment=TA_JUSTIFY)
        code_style = ParagraphStyle(name='Code', fontName='Courier', fontSize=8, leading=10)

        Story.append(Paragraph("VBA Code Documentation", title_style))
        Story.append(Spacer(1, 12))

        for module_name, module_data in analysis.items():
            Story.append(Paragraph(f"Module: {module_name}", heading1_style))
            Story.append(Spacer(1, 12))

            if module_data['functions']:
                Story.append(Paragraph("Functions and Subroutines:", heading2_style))
                func_data = [['Name', 'Parameters']] + [[func[0], func[1]] for func in module_data['functions']]
                t = Table(func_data, colWidths=[200, 300])
                t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                       ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                       ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                       ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                       ('FONTSIZE', (0, 0), (-1, 0), 14),
                                       ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                       ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                       ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                                       ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                       ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                                       ('FONTSIZE', (0, 1), (-1, -1), 12),
                                       ('TOPPADDING', (0, 1), (-1, -1), 6),
                                       ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
                                       ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
                Story.append(t)
                Story.append(Spacer(1, 12))

            if module_data['variables']:
                Story.append(Paragraph("Variables:", heading2_style))
                var_data = [['Name', 'Type']] + module_data['variables']
                t = Table(var_data, colWidths=[200, 300])
                t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                       ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                       ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                       ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                       ('FONTSIZE', (0, 0), (-1, 0), 14),
                                       ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                       ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                       ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                                       ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                       ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                                       ('FONTSIZE', (0, 1), (-1, -1), 12),
                                       ('TOPPADDING', (0, 1), (-1, -1), 6),
                                       ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
                                       ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
                Story.append(t)
                Story.append(Spacer(1, 12))

            if module_data['comments']:
                Story.append(Paragraph("Comments:", heading2_style))
                for comment in module_data['comments'][:10]: 
                    Story.append(Paragraph(comment, normal_style))
                if len(module_data['comments']) > 10:
                    Story.append(Paragraph(f"... and {len(module_data['comments']) - 10} more comments", normal_style))
                Story.append(Spacer(1, 12))

            if module_data['logic']:
                Story.append(Paragraph("Logic:", heading2_style))
                for logic in module_data['logic']:
                    Story.append(Paragraph(logic, normal_style))
                Story.append(Spacer(1, 12))

            Story.append(Paragraph("Full Code:", heading2_style))
            code_lines = module_data['full_code'].split('\n')
            for line in code_lines:
                Story.append(Paragraph(line, code_style))
            Story.append(Spacer(1, 12))

        doc.build(Story)
        
        return pdf_filename

    except Exception as e:
        raise RuntimeError(f"An error occurred while generating PDF documentation: {str(e)}")

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    try:
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)

