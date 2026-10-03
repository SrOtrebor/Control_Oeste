import sys
import re

filepath = 'public/js/nomina_logic_snippet.js'
with open(filepath, 'r', encoding='utf-8') as f:
    text = f.read()

# Change modal interactions
text = text.replace("modalNomina.classList.add('active');", "document.querySelector('button[data-tab=\"config\"]').click(); setTimeout(() => document.getElementById('formNomina').scrollIntoView({behavior: 'smooth', block: 'start'}), 100);")
text = text.replace("modalNomina.classList.remove('active');", "// no modal to remove")
text = text.replace("window.addEventListener('click', (e) => {\n    if (e.target == modalNomina) {\n        closeNomina();\n    }\n});", "")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(text)
print('Updated nomina_logic_snippet.js')
