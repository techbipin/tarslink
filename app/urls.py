from django.urls import path
from app.views import (
    index,
    process_file,
    generate_output,
    generate_graph
)

urlpatterns = [
    path('', index, name='index'),
    path('process-file/', process_file, name='process-file'),
    path('generate-output/', generate_output, name='generate-output'),
    path('generate-graph/', generate_graph, name='generate-graph'),
]
