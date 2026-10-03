import sys
import re

filepath = 'public/dashboard.html'
with open(filepath, 'r', encoding='utf-8') as f:
    text = f.read()

# Extract the formNomina form
form_match = re.search(r'(<form id=\"formNomina\">.*?</form>)', text, re.DOTALL)
if not form_match:
    print('formNomina not found')
    sys.exit(1)

form_html = form_match.group(1)

# Remove the entire modal from HTML
modal_start = text.find('<div id="modalNomina" class="modal">')
if modal_start != -1:
    modal_end = text.find('</form>', modal_start) + len('</form>')
    modal_end = text.find('</div>', modal_end)
    modal_end = text.find('</div>', modal_end + 1)
    modal_end = text.find('</div>', modal_end + 1) + len('</div>')
    text = text[:modal_start] + text[modal_end:]

# Wrap the form in a config card
card_html = f'''
                    <!-- ABM Personal (Ex Nómina Form) -->
                    <div class="config-card" id="configCardNomina" style="grid-column: 1 / -1; background: var(--dash-surface-hover); padding: 1.5rem; border-radius: var(--dash-radius); border: 1px solid var(--dash-border);">
                        <h3 style="margin-bottom: 1rem; color: var(--dash-primary); display: flex; align-items: center; gap: 0.5rem;">
                            <span class="material-icons">badge</span> Agregar / Editar Personal Autorizado
                        </h3>
                        {form_html}
                    </div>
'''

config_content_end = text.find('<!-- Respaldo de Base de Datos -->')
if config_content_end != -1:
    right_col_end = text.find('</div>', text.find('</div>', config_content_end) + 6) + 6
    text = text[:right_col_end] + card_html + text[right_col_end:]
else:
    print('Could not find where to insert')
    sys.exit(1)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(text)
print('Moved form to config tab')
