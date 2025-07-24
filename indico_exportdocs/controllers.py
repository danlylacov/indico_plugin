from flask import send_file, render_template
from indico.core.plugins import IndicoPluginBlueprint
from io import BytesIO
from .util import generate_docx_list, generate_docx_report, generate_docx_papers
from indico.modules.events.management.views import WPEventManagement
from indico.modules.events.management.controllers.base import RHManageEventBase
import os

blueprint = IndicoPluginBlueprint(
    'exportdocs',
    __name__,
    url_prefix='/event/<int:event_id>/manage'
)
@blueprint.route('/export/list')
def export_list(event_id):
    docx_bytes = generate_docx_list(event_id)
    return send_file(BytesIO(docx_bytes), as_attachment=True, download_name='list.docx')

@blueprint.route('/export/report')
def export_report(event_id):
    docx_bytes = generate_docx_report(event_id)
    return send_file(BytesIO(docx_bytes), as_attachment=True, download_name='report.docx')

@blueprint.route('/export/papers')
def export_papers(event_id):
    docx_bytes = generate_docx_papers(event_id)
    return send_file(BytesIO(docx_bytes), as_attachment=True, download_name='papers.docx')

class RHExportDocs(RHManageEventBase):
    print('DEBUG TEMPLATE PATH:', os.path.abspath(os.path.join(os.path.dirname(__file__), 'templates/exportdocs/buttons.html')))
    def _process(self):
        import os
        print('TEMPLATE EXISTS:', os.path.exists('/home/daniil/dev/indico/indico_exportdocs/templates/exportdocs/buttons.html'))
        return WPEventManagement.render_template('buttons.html', self.event)

blueprint.add_url_rule('/export', 'export_buttons', RHExportDocs)
