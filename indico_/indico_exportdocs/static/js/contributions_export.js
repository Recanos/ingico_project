/**
 * Добавляет кнопку экспорта на страницу управления докладами
 */
(function() {
    'use strict';

    function addExportButton() {
        // Проверяем, что мы на странице управления докладами
        if (!window.location.pathname.includes('/manage/contributions')) {
            return;
        }

        // Ищем различные места для добавления кнопки
        const locations = [
            // Панель действий в заголовке
            '.page-header .actions',
            // Панель инструментов
            '.toolbar',
            // Панель фильтров
            '.filtering-toolbar',
            // Верхняя панель
            '.top-actions',
            // Панель действий таблицы
            '.table-actions'
        ];

        let buttonAdded = false;

        for (const selector of locations) {
            const container = document.querySelector(selector);
            if (container && !buttonAdded) {
                // Создаем кнопку
                const button = document.createElement('a');
                button.href = window.location.pathname.replace('/contributions', '/export/list');
                button.className = 'i-button icon-file-word highlight';
                button.title = 'Экспорт списка докладов в DOCX';
                button.innerHTML = '<span>Печать отчета</span>';
                
                // Добавляем кнопку в начало контейнера
                container.insertBefore(button, container.firstChild);
                buttonAdded = true;
                console.log('Кнопка экспорта добавлена в:', selector);
                break;
            }
        }

        // Если не удалось найти стандартные места, добавляем в заголовок страницы
        if (!buttonAdded) {
            const pageHeader = document.querySelector('.page-header');
            if (pageHeader) {
                const actionsDiv = pageHeader.querySelector('.actions') || pageHeader;
                const button = document.createElement('a');
                button.href = window.location.pathname.replace('/contributions', '/export/list');
                button.className = 'i-button icon-file-word highlight';
                button.title = 'Экспорт списка докладов в DOCX';
                button.innerHTML = '<span>Печать отчета</span>';
                
                actionsDiv.appendChild(button);
                console.log('Кнопка экспорта добавлена в заголовок страницы');
            }
        }
    }

    // Запускаем после загрузки DOM
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', addExportButton);
    } else {
        addExportButton();
    }

    // Также запускаем после AJAX загрузок (для SPA)
    document.addEventListener('indico:pageLoaded', addExportButton);
    document.addEventListener('indico:contributionsLoaded', addExportButton);
})();
