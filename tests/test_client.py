# test_final.py
import sys
sys.path.insert(0, '.')

from src.excel_preview import generate_final_preview
import yaml

clients = yaml.safe_load(open('data/input/clients.yaml'))
df = generate_final_preview(clients)

# Show rows with client
with_client = df[df['client'] != '']
print(with_client[['week_beginning', 'category', 'client', 'hours', 'opportunity_id']].to_string())