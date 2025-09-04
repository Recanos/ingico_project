from indico.core.plugins import IndicoPlugin


class ExportDocsPlugin(IndicoPlugin):
    """Экспорт отчетов и списков в docx"""
    
    def get_blueprints(self):
        # Ленивый импорт для избежания циклических импортов
        from .controllers import blueprint
        return blueprint
    
    def get_assets(self):
        """Возвращает JavaScript и CSS файлы для плагина."""
        return {
            'js': ['js/contributions_export.js'],
            'css': []
        }
