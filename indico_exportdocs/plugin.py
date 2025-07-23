from indico.core.plugins import IndicoPlugin
from .controllers import blueprint

print('DEBUG plugin.py loaded')

class ExportDocsPlugin(IndicoPlugin):
    """Экспорт отчетов и списков в docx"""
    def get_blueprints(self):
        print('DEBUG get_blueprints:', type(blueprint), blueprint)
        print('DEBUG get_blueprints return:', (blueprint,))
        return blueprint
