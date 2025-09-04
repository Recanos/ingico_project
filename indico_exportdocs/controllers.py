from flask import send_file, render_template_string
from indico.core.plugins import IndicoPluginBlueprint
from io import BytesIO
from .util import generate_docx_list, generate_docx_report, generate_docx_papers
from indico.modules.events.management.controllers.base import RHManageEventBase

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
    """Контроллер для отображения страницы экспорта документов."""
    
    def _process(self):
        html = f'''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Экспорт документов - {self.event.title}</title>
            <style>
                body {{ 
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; 
                    margin: 0; 
                    padding: 20px; 
                    background: #f8f9fa; 
                    color: #333;
                }}
                .container {{ 
                    max-width: 900px; 
                    margin: 0 auto; 
                    background: white; 
                    border-radius: 8px; 
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    overflow: hidden;
                }}
                .header {{ 
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                    color: white; 
                    padding: 30px; 
                    text-align: center;
                }}
                .header h1 {{ margin: 0; font-size: 2.5em; font-weight: 300; }}
                .header p {{ margin: 10px 0 0 0; opacity: 0.9; font-size: 1.1em; }}
                
                .content {{ padding: 40px; }}
                
                .buttons {{ 
                    display: grid; 
                    grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                    gap: 20px; 
                    margin: 30px 0;
                }}
                .btn {{ 
                    display: block; 
                    padding: 20px; 
                    background: white; 
                    color: #333; 
                    text-decoration: none; 
                    border-radius: 8px; 
                    font-weight: 500;
                    text-align: center;
                    border: 2px solid #e9ecef;
                    transition: all 0.3s ease;
                    box-shadow: 0 2px 5px rgba(0,0,0,0.05);
                }}
                .btn:hover {{ 
                    transform: translateY(-2px); 
                    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
                    border-color: #667eea;
                }}
                .btn.highlight {{ 
                    background: linear-gradient(135deg, #28a745 0%, #20c997 100%); 
                    color: white; 
                    border-color: #28a745;
                }}
                .btn.highlight:hover {{ 
                    background: linear-gradient(135deg, #218838 0%, #1ea085 100%);
                }}
                
                .btn-icon {{ 
                    font-size: 2em; 
                    display: block; 
                    margin-bottom: 10px;
                }}
                .btn-title {{ 
                    font-size: 1.2em; 
                    font-weight: 600; 
                    margin-bottom: 5px;
                }}
                .btn-desc {{ 
                    font-size: 0.9em; 
                    opacity: 0.8;
                }}
                
                .info {{ 
                    background: #e7f3ff; 
                    padding: 25px; 
                    border-radius: 8px; 
                    margin: 30px 0;
                    border-left: 4px solid #007bff;
                }}
                .info h3 {{ margin: 0 0 15px 0; color: #0056b3; }}
                .info p {{ margin: 0; line-height: 1.6; }}
                
                .back-link {{ 
                    text-align: center; 
                    margin-top: 40px; 
                    padding-top: 20px; 
                    border-top: 1px solid #e9ecef;
                }}
                .back-link a {{ 
                    color: #667eea; 
                    text-decoration: none; 
                    font-weight: 500;
                    padding: 10px 20px;
                    border: 2px solid #667eea;
                    border-radius: 6px;
                    transition: all 0.3s ease;
                }}
                .back-link a:hover {{ 
                    background: #667eea; 
                    color: white;
                }}
                
                @media (max-width: 768px) {{
                    .buttons {{ grid-template-columns: 1fr; }}
                    .header h1 {{ font-size: 2em; }}
                    .content {{ padding: 20px; }}
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📄 Экспорт документов и публикаций</h1>
                    <p>Событие: <strong>{self.event.title}</strong></p>
                    <p style="font-size: 0.9em; opacity: 0.8;">Оформление по ГОСТ Р 7.0.97-2016</p>
                </div>
                
                <div class="content">
                    <div class="buttons">
                        <a href="/event/{self.event.id}/manage/export/list" class="btn highlight">
                            <span class="btn-icon">📋</span>
                            <div class="btn-title">Список докладов</div>
                            <div class="btn-desc">Таблица по дням: №, ФИО+название, Статус, Решение</div>
                        </a>
                        
                        <a href="/event/{self.event.id}/manage/export/report" class="btn">
                            <span class="btn-icon">📊</span>
                            <div class="btn-title">Отчет о проведении</div>
                            <div class="btn-desc">Пронумерованный список по дням: "1. ФИО. Название доклада"</div>
                        </a>
                        
                        <a href="/event/{self.event.id}/manage/export/papers" class="btn">
                            <span class="btn-icon">📝</span>
                            <div class="btn-title">Список публикаций</div>
                            <div class="btn-desc">Статьи по дням со статусом "приняты к публикации"</div>
                        </a>
                    </div>
                    
                    <div class="info">
                        <h3>ℹ️ Формат документов</h3>
                        <p>Каждая функция создает документ с разбивкой по времени докладов в соответствии с <strong>ГОСТ Р 7.0.97-2016</strong>:</p>
                        <ul style="margin: 15px 0; padding-left: 20px; line-height: 1.6;">
                            <li><strong>Список докладов</strong> - таблица для каждого дня с колонками: №, ФИО+название, Статус, Решение</li>
                            <li><strong>Отчет о проведении</strong> - пронумерованный список по дням: "1. ФИО. Название доклада"</li>
                            <li><strong>Список публикаций</strong> - статьи со статусом "приняты к публикации" по дням: "1. ФИО, группа. Название статьи"</li>
                        </ul>
                        <p><strong>Время берется из расписания (timetable)</strong> каждого доклада. Доклады группируются по датам.</p>
                        <p><strong>Оформление по ГОСТ:</strong> Times New Roman 14 пт, межстрочный интервал 1,5, поля 20/10/20/20 мм.</p>
                    </div>
                    
                    <div class="back-link">
                        <a href="/event/{self.event.id}/manage">← Вернуться к управлению событием</a>
                    </div>
                </div>
            </div>
        </body>
        </html>
        '''
        return html

#  маршрут для страницы экспорта
blueprint.add_url_rule('/export', 'export_buttons', RHExportDocs)
