from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION
from pathlib import Path

out = Path(r'C:\Users\User\Desktop\1030\outputs\個人行程與任務管理系統_PRD.docx')
doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(.72); sec.bottom_margin = Inches(.72); sec.left_margin = Inches(.78); sec.right_margin = Inches(.78)
styles=doc.styles
for name in ['Normal','Title','Heading 1','Heading 2','Heading 3']:
    st=styles[name]
    st.font.name='Microsoft JhengHei'; st._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
    st.font.color.rgb=RGBColor(0,0,0)
styles['Normal'].font.size=Pt(10); styles['Normal'].paragraph_format.space_after=Pt(5)
styles['Title'].font.size=Pt(25); styles['Title'].font.bold=True; styles['Title'].paragraph_format.space_after=Pt(6)
styles['Heading 1'].font.size=Pt(16); styles['Heading 1'].font.bold=True; styles['Heading 1'].paragraph_format.space_before=Pt(14); styles['Heading 1'].paragraph_format.space_after=Pt(6)
styles['Heading 2'].font.size=Pt(12); styles['Heading 2'].font.bold=True; styles['Heading 2'].paragraph_format.space_before=Pt(9); styles['Heading 2'].paragraph_format.space_after=Pt(4)

def shade(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)
def cell_margins(cell, top=100, start=110, bottom=100, end=110):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr(); mar=OxmlElement('w:tcMar')
    for m,v in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        e=OxmlElement('w:'+m); e.set(qn('w:w'),str(v)); e.set(qn('w:type'),'dxa'); mar.append(e)
    tcPr.append(mar)
def table(headers, rows, widths=None):
    t=doc.add_table(rows=1, cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
    for i,h in enumerate(headers):
        c=t.rows[0].cells[i]; c.text=h; shade(c,'6F6048'); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; cell_margins(c,130,120,130,120)
        for r in c.paragraphs[0].runs: r.font.bold=True; r.font.color.rgb=RGBColor(255,255,255); r.font.size=Pt(9)
    trPr=t.rows[0]._tr.get_or_add_trPr(); repeat=OxmlElement('w:tblHeader'); repeat.set(qn('w:val'),'true'); trPr.append(repeat)
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        for i,val in enumerate(row):
            cells[i].text=str(val); cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER; cell_margins(cells[i])
            if ri%2: shade(cells[i],'F5F2EB')
            for p in cells[i].paragraphs:
                p.paragraph_format.space_after=Pt(1)
                for run in p.runs: run.font.size=Pt(8.5)
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths): row.cells[i].width=Inches(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(2)
    return t

def para(text='', boldlead=None):
    p=doc.add_paragraph()
    if boldlead and text.startswith(boldlead): p.add_run(boldlead).bold=True; p.add_run(text[len(boldlead):])
    else: p.add_run(text)
    return p
def bullets(items):
    for x in items: doc.add_paragraph(x, style='List Bullet')
def numbered(items):
    for x in items: doc.add_paragraph(x, style='List Number')

p=doc.add_paragraph(style='Title'); p.add_run('個人行程與任務管理系統')
p=doc.add_paragraph(); r=p.add_run('產品需求文件 PRD'); r.bold=True; r.font.size=Pt(13); r.font.color.rgb=RGBColor(111,96,72)
para('版本 1.0｜需求討論初稿｜2026 年 10 月 30 日')
para('本文件定義一套跨桌機與手機使用的個人規劃產品，整合行程、待辦、習慣、專案進度與靈感記錄，並以日、週、月總覽支援規劃與執行。第一階段以響應式網頁原型驗證核心操作；行動 App 與桌面常駐便利貼屬正式產品方案，需於技術規劃階段決定實作形式。')

doc.add_heading('1. 產品概述', level=1)
para('產品暫名為「日常」，定位為個人使用的行程與任務管理系統。使用者可以在單一總覽中查看不同時間尺度的安排，並透過分頁或篩選切換資料類別；事項完成後可直接勾選，系統保留每日完成狀態，方便回顧當天完成情形。')
table(['項目','定義'],[
('產品目標','讓個人能快速規劃、查看並完成日常安排，減少在多個工具間切換。'),
('主要平台','桌機與手機皆支援；需求包含網頁與 App 形式。'),
('主要介面','共用總覽搭配分類分頁／篩選，提供日、週、月行事曆視圖。'),
('視覺方向','低刺激、柔和的大地色系；以暖米白、陶土、橄欖與沙岩色區分資訊。'),
('桌面便利貼','使用者希望像便利貼一樣在桌面顯示摘要；原型中的釘選按鈕僅為互動示意，非作業系統常駐功能。')], [1.35,5.65])

doc.add_heading('2. 目標使用者與需求', level=1)
para('目標使用者是需要自行管理個人時間、日常任務、習慣與多個進行中計畫的人。其核心需求包括：')
bullets(['以日／週／月瀏覽行程與事項內容，掌握安排密度。','管理待辦任務，完成時能勾選，並保留每日完成紀錄（不要求記錄完成的具體時間）。','設定目標或習慣，查看每日打卡及一週完成狀態。','追蹤專案進度，快速了解各專案目前完成比例。','記錄靈感，讓想法可以與日常管理集中於同一產品。','在手機與桌機使用，並能在桌面快速查看個人摘要。'])

doc.add_heading('3. 產品範圍', level=1)
doc.add_heading('3.1 第一階段：原型／MVP 核心', level=2)
bullets(['響應式總覽頁與日、週、月視圖切換。','行程、待辦、習慣打卡、專案進度與靈感的分類呈現。','新增事項及完成狀態切換。','每日完成狀態保存；原型階段以瀏覽器本機儲存示範。','桌機與手機版介面適配。','桌面便利貼的概念與視覺示意。'])
doc.add_heading('3.2 待後續評估／不納入目前原型承諾', level=2)
bullets(['正式 iOS／Android 原生或跨平台 App 的技術選型與上架。','真正置頂於桌面的 Widget、系統小工具或桌面 App。','帳號、雲端同步、多裝置資料一致性及離線衝突處理。','行事曆外部服務同步、通知提醒、分享與協作。','資料匯出、備份、權限與隱私政策細節。'])

doc.add_heading('4. 資訊架構與導覽', level=1)
table(['區域','內容／操作'],[
('總覽','摘要卡、行事曆、所選日期的事項清單、習慣週追蹤、專案進度、靈感角落。'),
('行程','檢視行程事項，支援日／週／月切換。'),
('待辦任務','檢視與勾選任務，查看完成狀態。'),
('目標與習慣','查看習慣清單及每週每日打卡狀態。'),
('專案','查看專案完成比例及進度摘要。'),
('靈感筆記','新增、瀏覽靈感內容。')], [1.5,5.5])
para('導覽以「共用總覽＋分類分頁／篩選」為原則。桌機可使用側邊導覽；手機以底部導覽呈現常用入口。分類切換應保留目前選取日期與視圖狀態，除非使用者明確返回或切換時間範圍。')

doc.add_heading('5. 核心使用流程', level=1)
numbered(['使用者進入總覽，預設看到目前日期附近的行程與每日摘要。','使用者透過日／週／月控制切換時間尺度，利用前後日期控制瀏覽範圍，或返回今天。','使用者點選日期，該日期成為目前選取日，下方清單與摘要隨之更新。','使用者在總覽新增行程、待辦、習慣或靈感，輸入名稱、類型、日期／時間及可選備註。','使用者在事項列勾選完成；系統更新該事項的完成狀態，並保留其所屬日期的完成紀錄，不記錄完成時間。','使用者在習慣週追蹤中點選日期，切換該日打卡狀態。','使用者檢視專案進度摘要，以了解各專案完成比例。'])

doc.add_heading('6. 功能需求', level=1)
features=[
('FR-01','時間視圖切換','使用者可在日、週、月視圖間切換；目前視圖需有明確選取狀態。','P0'),
('FR-02','日期瀏覽','可前往上一／下一日、週或月，並提供返回今天操作。','P0'),
('FR-03','行事曆內容','行事曆格應顯示事項名稱及可用時間，不得只以事項類型或色點呈現；內容過多時需提示尚有項目。','P0'),
('FR-04','日期選取','點選行事曆日期後，更新選取日期、摘要及事項清單。','P0'),
('FR-05','新增事項','可建立行程、待辦、習慣打卡或靈感；名稱為必填，備註及時間可選。','P0'),
('FR-06','完成勾選','行程／待辦等適用項目可切換完成狀態，完成狀態依所屬日期保存；不需完成時間戳記。','P0'),
('FR-07','每日完成紀錄','可辨識指定日期哪些事項已完成；切換日期後仍顯示該日的完成狀態。','P0'),
('FR-08','習慣追蹤','呈現每週習慣列與每日打卡狀態；使用者可切換打卡。','P0'),
('FR-09','專案進度','以完成比例與摘要呈現多個專案。MVP 原型使用示範資料；正式版需定義進度編輯方式。','P1'),
('FR-10','靈感記錄','可建立及呈現靈感筆記，並依日期納入清單或分類檢視。','P1'),
('FR-11','分類篩選','可依全部、行程、待辦、習慣及靈感篩選所選日期的清單。','P0'),
('FR-12','本機持久化','原型可在同一瀏覽器保存資料；正式版需另定帳號與雲端同步需求。','P0 原型'),
('FR-13','桌面便利貼','提供桌面摘要面板之設計方向；正式實作需確認作業系統、更新頻率及互動範圍。','待決')]
table(['ID','功能','需求描述','優先級'],features,[.62,1.25,4.45,.75])

doc.add_heading('7. 資料需求（初步）', level=1)
table(['資料實體','建議欄位','說明'],[
('事項 Item','ID、標題、類型、日期、時間（選填）、備註（選填）、完成狀態、標籤／分類、顯示色','類型至少包含行程、待辦、習慣、靈感。完成狀態需能按日期回顧；不要求完成時間。'),
('習慣 Habit','ID、名稱、圖示（選填）、目標頻率（待定）、建立狀態','打卡資料應獨立按日期保存，支援歷史檢視與連續天數計算（連續天數是否必要待確認）。'),
('習慣打卡 HabitLog','習慣 ID、日期、是否完成','每日一筆狀態；重複點選可撤銷或重新打卡。'),
('專案 Project','ID、名稱、總任務數、已完成數、目標日期（選填）、狀態','進度比例 = 已完成數 ÷ 總數；正式版需決定是否由任務自動計算。'),
('靈感 Inspiration','ID、內容、建立日期、標籤（選填）','現階段原型以一般事項呈現；是否需要長文、附件或搜尋待確認。')], [1.35,2.65,3.0])

doc.add_heading('8. 介面與互動規範', level=1)
bullets(['視覺：採柔和大地色系，背景暖米白、文字深褐灰，橄欖綠作為主要互動與完成色；分類色彩須低飽和且具足夠辨識度。','行事曆：週視圖顯示每格最多 3 項、月視圖最多 2 項並提供「＋N 項」提示；具體數量可依實際可用高度調整。','完成狀態：以勾選框、文字弱化或刪線等多重方式呈現，不能只靠顏色區別。','響應式：桌機採側邊導覽及多欄摘要；窄螢幕採單欄內容與底部導覽，行事曆事項仍須呈現可辨識的內容文字。','新增流程：以清楚標示必填欄位、類型與可選欄位的表單呈現；新增後立即更新畫面。','空狀態：無事項時提供友善說明與新增入口；錯誤狀態需保留使用者已輸入內容。'])

doc.add_heading('9. 非功能需求', level=1)
table(['面向','需求'],[
('易用性','主要動作（切換視圖、選日期、勾選、新增）可直接從總覽完成，避免不必要的多層導覽。'),
('響應式與可及性','適配桌機及手機寬度；互動元件具足夠點擊尺寸、鍵盤焦點與可讀標籤。資訊狀態不只依賴顏色。'),
('效能','一般個人規模資料的總覽切換與勾選應即時回饋；載入大量歷史資料時需採分頁或區間載入（正式版評估）。'),
('資料可靠性','正式版需明確定義本機快取、雲端同步、備份、刪除復原及多裝置衝突處理策略。'),
('隱私與安全','行程與靈感視為個人資料；登入、傳輸加密、資料保存及刪除政策於正式產品設計階段補齊。')], [1.35,5.65])

doc.add_heading('10. 驗收條件', level=1)
numbered(['使用者可切換日、週、月視圖，並可前後瀏覽及返回今天。','選取某日期後，該日期的事項與完成狀態正確顯示；切換離開再回來，完成狀態仍保留。','行事曆顯示事項名稱，非僅顯示類別圖示或色點；超出顯示上限時有剩餘數量提示。','使用者可新增至少一筆行程、待辦、習慣及靈感事項，新增後立即出現在正確日期與篩選結果。','勾選與取消勾選均能更新完成樣式與摘要，且不必輸入或顯示完成時間。','習慣打卡可以按日切換，並在週追蹤中正確反映。','在常見手機及桌機 viewport 下，核心操作可用，行事曆文字不被隱藏成純色點。','專案摘要能呈現名稱、完成比例及簡短進度資訊。'])

doc.add_heading('11. 衡量指標（建議）', level=1)
para('以下為產品上線後可評估的指標，需待目標族群與分析方案確認後設定基準與目標值。')
bullets(['新使用者首次建立事項的完成率與耗時。','每週活躍使用者建立／完成事項的比例。','習慣打卡週完成率與連續使用週數。','日／週／月視圖使用分布及跨視圖切換率。','多裝置使用比例及同步失敗率（正式版）。'])

doc.add_heading('12. 假設、限制與待確認事項', level=1)
para('本 PRD 根據目前已確認的需求整理。尚未確認的產品政策以待決事項標示，避免將原型中的示範行為誤當成正式規格。')
table(['待確認議題','影響／建議'],[
('App 形式','確認是否先以 PWA 提供可安裝體驗，或需 iOS／Android 原生或跨平台 App。建議先確認離線、推播及系統行事曆整合是否為必要條件。'),
('桌面便利貼','確認目標作業系統（Windows／macOS）、是否需真正置頂、是否可直接打勾，以及背景同步需求。'),
('帳號與同步','是否登入、跨裝置同步、共用或家庭成員支援。若多裝置為必要情境，建議列入 MVP 架構設計。'),
('提醒與重複規則','是否需要行程提醒、到期通知、重複行程或重複任務。'),
('任務與專案關聯','任務是否可歸屬專案；專案進度是否依關聯任務完成狀態自動計算。'),
('習慣設定','習慣目標頻率、補打卡／撤銷策略、連續天數與歷史視圖規則。'),
('靈感筆記','是否需要長文、標籤、搜尋、附件或將靈感轉成任務。'),
('日期與時間','預設時區、全天事項、跨日事項及沒有時間的事項如何排序。')], [1.55,5.45])

doc.add_heading('附錄 A 原型對應', level=1)
para('目前互動原型檔案為 outputs/index.html，包含示範資料、日／週／月視圖、內容顯示、事項勾選、新增表單、習慣週追蹤、專案進度及靈感區。原型資料以瀏覽器 localStorage 保存；「桌面便利貼」控制僅模擬釘選狀態。')

# footer page numbering field
for section in doc.sections:
    f=section.footer.paragraphs[0]; f.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=f.add_run('個人行程與任務管理系統 PRD  ·  '); r.font.size=Pt(8); r.font.color.rgb=RGBColor(130,124,112)
    fld=OxmlElement('w:fldSimple'); fld.set(qn('w:instr'),'PAGE'); f._p.append(fld)

doc.core_properties.title='個人行程與任務管理系統 PRD'
doc.core_properties.subject='產品需求文件'
doc.core_properties.author=''
doc.save(out)
print(out)
