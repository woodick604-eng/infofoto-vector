
import os

with open('/Users/alex/Desktop/INFOFOTO VECTOR/container/static/logo_new_b64.txt', 'r') as f:
    full_b64 = f.read().strip()

app_path = '/Users/alex/Desktop/INFOFOTO VECTOR/container/app.py'
with open(app_path, 'r') as f:
    content = f.read()

# Busquem la constant LOGO_PMA_B64 i la substituïm completament
start_marker = 'LOGO_PMA_B64 = """'
end_marker = '"""'
start_idx = content.find(start_marker)
if start_idx != -1:
    end_idx = content.find(end_marker, start_idx + len(start_marker))
    if end_idx != -1:
        new_constant = f'LOGO_PMA_B64 = """{full_b64}"""'
        new_content = content[:start_idx] + new_constant + content[end_idx + len(end_marker):]
        with open(app_path, 'w') as f:
            f.write(new_content)
        print("PATCH: LOGO_PMA_B64 actualitzat amb el logo de Mossos + Generalitat.")
    else:
        print("ERROR: No s'ha trobat el final de la constant.")
else:
    print("ERROR: No s'ha trobat la constant LOGO_PMA_B64.")
