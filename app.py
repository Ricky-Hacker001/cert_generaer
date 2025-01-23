from flask import Flask, render_template, request, redirect, url_for, flash
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
import pythoncom

app = Flask(__name__)
app.secret_key = 'your_secret_key'
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

# Ensure the upload and output folders exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def sanitize_filename(filename):
    return "".join([c if c.isalnum() or c in ' ._-()' else '_' for c in filename])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/form', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        count = int(request.form['count'])
        template_file = request.files['template']
        template_path = os.path.join(UPLOAD_FOLDER, template_file.filename)
        template_file.save(template_path)

        students_data = []
        for i in range(count):
            student_data = {
                'student_name': request.form[f'student_name_{i}'],
                'course': request.form[f'course_{i}'],
                'college_name': request.form[f'college_name_{i}'],
                'college_location': request.form[f'college_location_{i}'],
                'internship_domain': request.form[f'internship_domain_{i}'],
                'start_date': request.form[f'start_date_{i}'],
                'end_date': request.form[f'end_date_{i}'],
                'print_date': request.form[f'print_date_{i}']
            }
            students_data.append(student_data)

        for i, student_data in enumerate(students_data):
            doc = DocxTemplate(template_path)
            doc.render(student_data)
            sanitized_student_name = sanitize_filename(student_data['student_name'])
            sanitized_college_name = sanitize_filename(student_data['college_name'])
            output_path_docx = os.path.join(OUTPUT_FOLDER, f"{sanitized_student_name}_{sanitized_college_name}_Certificate_{i+1}.docx")
            doc.save(output_path_docx)

            # Initialize and Uninitialize COM
            pythoncom.CoInitialize()
            try:
                convert(output_path_docx)
            finally:
                pythoncom.CoUninitialize()

        flash(f"Generated {count} certificates successfully.")
        return redirect(url_for('index'))

    return render_template('form.html')

if __name__ == '__main__':
    app.run(debug=True)
