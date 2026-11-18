"""Generate a clean, generic 15-shot Excel template. Never overwrite existing assets."""
from pathlib import Path
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

def main():
    assets=Path(__file__).resolve().parent.parent/'assets'
    assets.mkdir(exist_ok=True)
    blank=assets/'blank.jpg'
    if not blank.exists():
        Image.new('RGB',(144,224),'white').save(blank,'JPEG',quality=90)
    dest=assets/'default-template.xlsx'
    if dest.exists():
        print(f'Already exists: {dest}')
        return
    wb=Workbook()
    ws=wb.active
    ws.title='Sheet1'
    ws.merge_cells('A1:K1')
    ws['A1']='信息流视频分镜脚本'
    ws['A2']='视频参数'
    ws.merge_cells('C2:K2')
    ws['C2']=''
    headers=['镜号','镜头类型','画面','示意图','景别','拍摄方法','文本','秒数','拍摄道具','道具图片','拍摄场景']
    for i,h in enumerate(headers,1):
        ws.cell(3,i,h)
    widths=[8,16,42,24,13,22,42,10,21,16,20]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.row_dimensions[1].height=32
    ws.row_dimensions[2].height=32
    ws.row_dimensions[3].height=26
    line=Side(style='thin',color='CBD5E1')
    for row in ws.iter_rows(min_row=1,max_row=18,min_col=1,max_col=11):
        for cell in row:
            cell.font=Font(name='Microsoft YaHei',size=10,color='172B4D')
            cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            cell.border=Border(left=line,right=line,top=line,bottom=line)
            if cell.row==3:
                cell.fill=PatternFill('solid',fgColor='DFEBF5')
                cell.font=Font(name='Microsoft YaHei',size=11,bold=True,color='172B4D')
    ws['A1'].font=Font(name='Microsoft YaHei',size=16,bold=True,color='FFFFFF')
    ws['A1'].fill=PatternFill('solid',fgColor='253A53')
    for r in range(4,19):
        ws.row_dimensions[r].height=172
        ws.cell(r,1,r-3)
        for c in [3,7]:
            ws.cell(r,c).alignment=Alignment(horizontal='left',vertical='center',wrap_text=True)
        for c,w,h in [('D',144,224),('J',96,150)]:
            im=XLImage(str(blank))
            im.width=w
            im.height=h
            ws.add_image(im,f'{c}{r}')
    ws.freeze_panes='C4'
    ws.sheet_view.showGridLines=False
    ws.print_options.horizontalCentered=True
    ws.print_title_rows='1:3'
    ws.print_area='A1:K18'
    wb.save(dest)
    print(dest)

if __name__=='__main__':
    main()
