from indico.core import signals
from indico.web.menu import SideMenuItem


@signals.menu.items.connect_via('event-management-sidemenu')
def _extend_event_management_menu(sender, event, **kwargs):
    """Добавляет пункт меню 'Экспорт документов' в боковое меню управления событиями."""
    # Упрощаем проверку - убираем session.user
    return SideMenuItem('exportdocs', 'Экспорт документов', 
                       f'/event/{event.id}/manage/export', 
                       section='reports', weight=10, icon='file-word')


def _inject_export_button(event, **kwargs):
    """Добавляет кнопку экспорта в правую часть заголовка управления событиями."""
    return '''
    <div class="group">
        <a href="/event/{}/manage/export/list" 
           class="i-button icon-file-word highlight"
           title="Экспорт списка докладов в DOCX">
            <span>Печать отчета</span>
        </a>
    </div>
    '''.format(event.id)


# Регистрируем хук шаблона
from indico.web.flask.templating import register_template_hook
register_template_hook('event-management-header-right', _inject_export_button)



