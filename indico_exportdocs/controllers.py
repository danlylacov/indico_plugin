from flask import send_file
from indico.core.plugins import IndicoPluginBlueprint
from io import BytesIO
from .util import generate_docx_list, generate_docx_report, generate_docx_papers

blueprint = IndicoPluginBlueprint(
    'exportdocs',
    __name__,
    url_prefix='/event/<int:event_id>'
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
