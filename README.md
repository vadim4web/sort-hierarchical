# BIM QA/QC – Hierarchical Sort Engine v1.5

---

## 1. Overview | Огляд

EN:
Google Apps Script tool for **hierarchical sorting of BIM QA/QC tables** with full formatting preservation.

UA:
Інструмент Google Apps Script для **ієрархічного сортування BIM QA/QC таблиць зі збереженням форматування**.

---

## 2. Core Idea | Основна ідея

EN:
The tool does NOT just sort rows — it builds hierarchy first, then sorts inside it.

UA:
Інструмент не просто сортує рядки — він спочатку будує ієрархію, потім сортує всередині неї.

---

## 3. Sorting Logic | Логіка сортування

EN:
Two-stage sorting model:

1. ROOT grouping (first token of column B)
2. Hierarchical sort inside group

UA:
Двохетапна модель:

1. Групування по ROOT (перший елемент колонки B)
2. Ієрархічне сортування всередині групи

---

## 4. Sort Order | Порядок

EN:
B → C → D → E → L → M

UA:
B → C → D → E → L → M

---

## 5. What it does | Функціонал

EN:
- Hierarchical BIM sorting (Revit-like logic)
- Preserves formatting (colors, fonts, styles)
- Handles numeric + text comparison
- Stable grouping by ROOT identifier
- Works directly in Google Sheets

UA:
- Ієрархічне BIM сортування (логіка як у Revit)
- Збереження форматування (кольори, стилі, шрифти)
- Обробка чисел і тексту
- Стабільне групування по ROOT
- Робота прямо в Google Sheets

---

## 6. Installation | Встановлення

EN:
1. Open Google Sheets  
2. Extensions → Apps Script  
3. Paste `sortHierarchical.js`  
4. Save project  
5. Run `sortHierarchical()`  

UA:
1. Відкрити Google Sheets  
2. Extensions → Apps Script  
3. Вставити `sortHierarchical.js`  
4. Зберегти  
5. Запустити `sortHierarchical()`  

---

## 7. Optional UI Version | Опціональний UI

EN:
Can be deployed as a lightweight web interface using Apps Script Web App.

UA:
Можна використовувати як Web App з простим UI інтерфейсом.

---

## 8. Key Insight | Ключова ідея

EN:
Sorting = Group first → Sort inside groups

UA:
Сортування = Спочатку групування → Потім сортування всередині

---

## 9. Notes | Примітки

EN:
Built for BIM QA/QC workflows in engineering production environments.

UA:
Розроблено для BIM QA/QC процесів у інженерному середовищі.

(vadym.c@b-eng-s.com)[https://vadim4web.github.io/]