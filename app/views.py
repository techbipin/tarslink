import os
import base64
from io import BytesIO
from app.models import Entry
from django.conf import settings
import plotly.graph_objects as go
from django.shortcuts import render, redirect
from app.process import process_input_data, generate_excel_output


def index(request):
    return render(request, 'index.html')

def process_file(request):
    if request.FILES and request.POST:
        task_workbook = request.FILES.get('excel-file', None)
        master_workbook = os.path.join(settings.BASE_DIR, 'master.xlsx')
        process_input_data(task_workbook=task_workbook, master_workbook=master_workbook)
    return redirect('index')

def generate_output(request):
    result = generate_excel_output()
    return result

def generate_graph(request):
    entry_data = Entry.objects.filter(year="2023", clubbed_name="Acko General").order_by('value')
    product_list = [each.product for each in entry_data]
    value_list = sorted([float(each.value) for each in entry_data])

    fig = go.Figure(
        data=[go.Bar(x=product_list, y=value_list)],
        layout_title_text="Acko General"
    )

    buffer = BytesIO()
    fig.write_image(buffer, format="png")
    buffer.seek(0)
    encoded_image = base64.b64encode(buffer.read()).decode("utf-8")

    return render(request, 'index.html', {"bar_plot": encoded_image})