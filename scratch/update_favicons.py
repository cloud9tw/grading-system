
import os
import re

templates_dir = 'templates'
favicon_link = "\n    <link rel=\"icon\" type=\"image/png\" href=\"{{ url_for('static', filename='favicon.png') }}\">"

for filename in os.listdir(templates_dir):
    if filename.endswith('.html'):
        path = os.path.join(templates_dir, filename)
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        if '<link rel="icon"' not in content:
            # Insert after <head> or <meta charset...>
            if '<head>' in content:
                content = content.replace('<head>', f'<head>{favicon_link}')
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(content)
                print(f"Updated {filename}")
