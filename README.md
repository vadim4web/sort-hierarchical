# BIM QA/QC – Hierarchical Sort Engine v2.0

---

## 1. Overview | Огляд

EN:
A Google Apps Script tool for hierarchical sorting of BIM QA/QC tables.

UA:
Інструмент Google Apps Script для багаторівневого сортування BIM QA/QC таблиць.

---

## 2. UI Version | Версія інтерфейсу

EN:
This version includes a lightweight Bootstrap-based web interface.

UA:
Ця версія містить легкий Bootstrap інтерфейс.

---

## 3. Sorting Logic | Логіка

EN:
Sorting priority:
B → C → D → E → L → M

UA:
Пріоритет сортування:
B → C → D → E → L → M

---

## 4. What it does | Що робить

EN:
- Sorts BIM tables hierarchically  
- Preserves formatting  
- Supports optional file/page grouping  

UA:
- Сортує BIM таблиці по ієрархії  
- Зберігає форматування  
- Підтримує групування по файлу і сторінці  

---

## 5. Installation (Google Apps Script) | Встановлення

EN:
1. Open Google Sheets  
2. Go to Extensions → Apps Script  
3. Create new project  
4. Paste `sortHierarchical.js`  
5. Save project  
6. Run function `sortHierarchical()`  

UA:
1. Відкрити Google Sheets  
2. Extensions → Apps Script  
3. Створити новий проект  
4. Вставити `sortHierarchical.js`  
5. Зберегти  
6. Запустити `sortHierarchical()`  

---

## 6. Optional Web Deploy | Опціональний веб-деплой

EN:
You can deploy this script as a Web App for UI access:
- Apps Script → Deploy → New deployment  
- Select “Web App”  
- Set access to “Anyone in workspace” or “Anyone with link”

UA:
Можна задеплоїти як Web App:
- Deploy → New deployment  
- Web App  
- Доступ: workspace або по лінку  

---

## 7. Usage | Використання

EN:
Run directly from Apps Script or UI button.

UA:
Запуск через Apps Script або UI кнопку.

---

## 8. Notes | Примітки

EN:
Designed for QA/QC workflows in BIM/VDC environments.

UA:
Розроблено для QA/QC BIM/VDC процесів.