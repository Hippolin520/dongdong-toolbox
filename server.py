import io, os, re, shutil, subprocess, tempfile, zipfile
from pathlib import Path
from contextlib import asynccontextmanager
import fitz
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches as PInches
from PIL import Image

ROOT=Path(__file__).resolve().parent
app=FastAPI(title='咚咚文件工具箱')
MAX=40*1024*1024
EXT={0:{'.doc','.docx'},1:{'.doc','.docx'},2:{'.xls','.xlsx'},3:{'.xls','.xlsx'},4:{'.xls','.xlsx'},5:{'.xls','.xlsx'},6:{'.ppt','.pptx'},**{i:{'.pdf'} for i in range(7,19)}}

def safe_name(s):return re.sub(r'[^\w.\-\u4e00-\u9fff]+','_',Path(s).name)[:100]
def lo(src,out,fmt):
    profile=out/'profile';profile.mkdir(exist_ok=True)
    cmd=['libreoffice',f'-env:UserInstallation=file://{profile}','--headless','--convert-to',fmt,'--outdir',str(out),str(src)]
    try:p=subprocess.run(cmd,capture_output=True,text=True,timeout=100)
    except subprocess.TimeoutExpired:raise ValueError('轉換逾時')
    result=out/(src.stem+'.'+fmt.split(':')[0])
    if not result.exists():raise ValueError('LibreOffice 轉換失敗：'+(p.stderr or p.stdout)[-350:])
    return result

def pdf_doc(src,password=''):
    d=fitz.open(src)
    if d.needs_pass and not d.authenticate(password):
        d.close();raise ValueError('PDF 已加密，請輸入正確密碼')
    return d

def render_pdf(src,out,ext='jpg',password=''):
    d=pdf_doc(src,password);paths=[]
    try:
        if len(d)>100:raise ValueError('最多處理 100 頁')
        for n,page in enumerate(d):
            pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
            p=out/f'{src.stem}_第{n+1}頁.{ext}'
            pix.save(str(p),output='jpeg' if ext=='jpg' else 'png');paths.append(p)
    finally:d.close()
    return paths

def bundle(paths,out,name):
    if len(paths)==1:return paths[0]
    dest=out/name
    with zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as z:
        for p in paths:z.write(p,p.name)
    return dest

def office_pdf(src,out):return lo(src,out,'pdf')

def pdf_to_word(src,out,password):
    d=pdf_doc(src,password);doc=Document()
    for i,page in enumerate(d):
        if i:doc.add_page_break()
        text=page.get_text('text')
        for line in text.splitlines():doc.add_paragraph(line)
    d.close();p=out/(src.stem+'.docx');doc.save(p);return p

def pdf_to_excel(src,out,password,mode='table'):
    """Reconstruct editable table geometry using detected cell rectangles.

    PDF lacks original workbook semantics; this is a best-effort reconstruction.
    """
    if mode not in ('table','visual','precision'):
        raise ValueError('未知 Excel 轉換模式')
    d=pdf_doc(src,password);w=Workbook();w.remove(w.active)
    try:
        if len(d)>100:raise ValueError('最多處理 100 頁')
        for i,page in enumerate(d):
            sheet=w.create_sheet(f'第{i+1}頁')
            sheet.sheet_view.showGridLines=False
            if mode=='visual':
                pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
                image_path=out/f'page_{i+1}.png';pix.save(str(image_path))
                pic=ExcelImage(str(image_path));pic.width=pix.width;pic.height=pix.height
                sheet.add_image(pic,'A1')
                sheet.column_dimensions['A'].width=page.rect.width/7
                sheet.sheet_properties.pageSetUpPr.fitToPage=True
                sheet.page_setup.fitToWidth=1;sheet.page_setup.fitToHeight=1
                continue
            try:
                found=page.find_tables(strategy='lines' if mode=='precision' else 'lines_strict')
                tables=list(found.tables)
                if not tables and mode=='precision':tables=list(page.find_tables(strategy='text').tables)
            except Exception:tables=[]
            # Create a coordinate grid based on actual PDF cell edges rather than
            # rounding each table cell to a coarse fixed-size grid.
            xedges=[0.0,float(page.rect.width)]
            yedges=[0.0,float(page.rect.height)]
            for t in tables:
                for rect in (t.cells or []):
                    if rect:
                        x0,y0,x1,y1=rect;xedges.extend((x0,x1));yedges.extend((y0,y1))
            # Include a modest regular grid for non-table text; avoid huge sheets.
            xedges.extend(j*36.0 for j in range(1,int(page.rect.width/36)+1))
            yedges.extend(j*16.0 for j in range(1,int(page.rect.height/16)+1))
            def edges(values,tol=2.0):
                values=sorted(max(0.0,float(v)) for v in values)
                groups=[]
                for v in values:
                    if groups and v-groups[-1][-1]<=tol:groups[-1].append(v)
                    else:groups.append([v])
                return [sum(g)/len(g) for g in groups]
            xs=edges(xedges);ys=edges(yedges)
            if len(xs)>250 or len(ys)>700:raise ValueError('PDF 表格結構過於複雜，請改用原始版面模式')
            from bisect import bisect_left
            def idx(values,pos):return max(1,min(len(values)-1,bisect_left(values,float(pos)+1.0)))
            for col in range(1,len(xs)):
                sheet.column_dimensions[get_column_letter(col)].width=max(1.5,min(70,(xs[col]-xs[col-1])/7.0))
            for row in range(1,len(ys)):
                sheet.row_dimensions[row].height=max(7,min(180,(ys[row]-ys[row-1])*0.75))
            occupied=set()
            def place(rr,cc,value):
                if value is None or not str(value).strip():return None
                if (rr,cc) in occupied:return None
                occupied.add((rr,cc));cell=sheet.cell(rr,cc,str(value)[:32767]);cell.alignment=Alignment(wrap_text=True,vertical='center')
                return cell
            for t in tables:
                values=t.extract();cells=t.cells or []
                for r,values_row in enumerate(values):
                    for c,value in enumerate(values_row):
                        n=r*t.col_count+c
                        if n>=len(cells) or not cells[n]:continue
                        x0,y0,x1,y1=cells[n]
                        rr=idx(ys,y0);cc=idx(xs,x0)
                        rr2=max(rr,idx(ys,y1-2));cc2=max(cc,idx(xs,x1-2))
                        cell=place(rr,cc,value)
                        if cell is None:continue
                        cell.border=Border(left=Side(style='thin',color='999999'),right=Side(style='thin',color='999999'),top=Side(style='thin',color='999999'),bottom=Side(style='thin',color='999999'))
                        # Merge only if region is unoccupied. PDF merged cells are
                        # inferred from geometry, not recovered from metadata.
                        if rr2>rr or cc2>cc:
                            region={(a,b) for a in range(rr,rr2+1) for b in range(cc,cc2+1)}
                            if not (region & (occupied-{(rr,cc)})):
                                try:sheet.merge_cells(start_row=rr,start_column=cc,end_row=rr2,end_column=cc2);occupied.update(region)
                                except ValueError:pass
            for block in page.get_text('dict',sort=True)['blocks']:
                for line in block.get('lines',[]):
                    spans=[sp for sp in line['spans'] if sp['text'].strip()]
                    if not spans:continue
                    x=min(sp['bbox'][0] for sp in spans);y=min(sp['bbox'][1] for sp in spans)
                    if any(fitz.Rect(t.bbox).contains(fitz.Point(x,y+2)) for t in tables):continue
                    rr=idx(ys,y);cc=idx(xs,x)
                    while (rr,cc) in occupied and cc<len(xs)-1:cc+=1
                    cell=place(rr,cc,''.join(sp['text'] for sp in spans).strip())
                    if cell:
                        sp=spans[0];rgb=sp.get('color',0)
                        cell.font=Font(name='Arial',size=max(7,min(22,sp.get('size',11)*0.75)),bold=bool(sp.get('flags',0)&16),color=f'{rgb:06X}')
                        cell.alignment=Alignment(vertical='center',wrap_text=False)
            sheet.sheet_properties.pageSetUpPr.fitToPage=True
            sheet.page_setup.fitToWidth=1
        result=out/(src.stem+'.xlsx');w.save(result);return result
    finally:d.close()

def pdf_to_ppt(src,out,password):
    d=pdf_doc(src,password);prs=Presentation();prs.slide_width=PInches(10);prs.slide_height=PInches(7.5)
    for page in d:
        slide=prs.slides.add_slide(prs.slide_layouts[6]);pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
        img=io.BytesIO(pix.tobytes('png'));w,h=page.rect.width,page.rect.height
        ratio=min(prs.slide_width/w,prs.slide_height/h);pw=int(w*ratio);ph=int(h*ratio)
        slide.shapes.add_picture(img,(prs.slide_width-pw)//2,(prs.slide_height-ph)//2,width=pw,height=ph)
    d.close();p=out/(src.stem+'.pptx');prs.save(p);return p

def split_excel(src,out,as_pdf):
    # LibreOffice can import legacy .xls, normalize it first.
    if src.suffix.lower()=='.xls':src=lo(src,out,'xlsx')
    wb=load_workbook(src,read_only=False,data_only=False)
    if len(wb.sheetnames)>100:raise ValueError('最多處理 100 張工作表')
    paths=[]
    for idx,sheet in enumerate(wb.worksheets,1):
        single=Workbook();target=single.active;target.title=sheet.title
        for row in sheet:
            for c in row:
                if c.value is not None:target.cell(c.row,c.column,c.value)
        for key,dim in sheet.column_dimensions.items():target.column_dimensions[key].width=dim.width
        for key,dim in sheet.row_dimensions.items():target.row_dimensions[key].height=dim.height
        target.sheet_format=sheet.sheet_format
        p=out/f'{safe_name(src.stem)}_{idx}_{safe_name(sheet.title)}.xlsx';single.save(p)
        if as_pdf:
            converted=office_pdf(p,out);paths.append(converted)
        else:paths.append(p)
    wb.close();return bundle(paths,out,src.stem+('_工作表PDF.zip' if as_pdf else '_工作表.zip'))

def process(index,files,out,option,password):
    src=files[0];stem=src.stem
    if index in (0,2,6):return office_pdf(src,out)
    if index in (1,3):return bundle(render_pdf(office_pdf(src,out),out,'jpg'),out,stem+'_圖片.zip')
    if index in (4,5):return split_excel(src,out,index==4)
    if index==7:return pdf_to_ppt(src,out,password)
    if index==12:return pdf_to_word(src,out,password)
    if index==13:return pdf_to_excel(src,out,password,option or 'table')
    if index in (14,15):return bundle(render_pdf(src,out,'jpg' if index==14 else 'png',password),out,stem+'_圖片.zip')
    if index==8:
        if len(files)<2:raise ValueError('至少需要兩個 PDF')
        result=fitz.open()
        for f in files:
            d=pdf_doc(f,password);result.insert_pdf(d);d.close()
        p=out/'合併文件.pdf';result.save(p,garbage=4,deflate=True);result.close();return p
    d=pdf_doc(src,password)
    try:
        if index==9:
            if not d.needs_pass:raise ValueError('這份 PDF 沒有密碼保護')
            p=out/(stem+'_解鎖.pdf');d.save(p,encryption=fitz.PDF_ENCRYPT_NONE,garbage=4);return p
        if index==10:
            angle=int(option or '90')
            if angle not in (90,180,270):raise ValueError('旋轉角度不正確')
            for page in d:page.set_rotation((page.rotation+angle)%360)
        elif index in (11,16):
            indices=parse_pages(option,len(d))
            if index==11:
                paths=[]
                for i in indices:
                    part=fitz.open();part.insert_pdf(d,from_page=i,to_page=i)
                    p=out/f'{stem}_第{i+1}頁.pdf';part.save(p);part.close();paths.append(p)
                return bundle(paths,out,stem+'_拆分.zip')
            if len(indices)==len(d):raise ValueError('不可刪除全部頁面')
            d.delete_pages(indices)
        elif index==17:
            for page in d:
                for widget in list(page.widgets() or []):
                    try:widget.update();page.delete_widget(widget)
                    except Exception:pass
                # rasterize each page to truly flatten visual content and annotations
            p=out/(stem+'_壓平.pdf');flat=fitz.open()
            for page in d:
                pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False,annots=True)
                rect=page.rect;new=flat.new_page(width=rect.width,height=rect.height)
                new.insert_image(rect,stream=pix.tobytes('png'))
            flat.save(p,garbage=4,deflate=True);flat.close();return p
        elif index==18:pass # rewrite cross-reference, streams and object structure
        else:raise ValueError('未知工具')
        p=out/(stem+{10:'_旋轉',16:'_刪除頁面',18:'_修復'}.get(index,'_處理')+'.pdf')
        d.save(p,garbage=4,deflate=True,clean=True);return p
    finally:d.close()

def parse_pages(value,n):
    if not value.strip():raise ValueError('請輸入頁碼')
    indices=set()
    for token in value.split(','):
        m=re.fullmatch(r'\s*(\d+)\s*(?:-\s*(\d+)\s*)?',token)
        if not m:raise ValueError('頁碼格式應為 1-3,5')
        a=int(m[1]);b=int(m[2] or a)
        if a<1 or b>n or a>b:raise ValueError(f'頁碼必須在 1 至 {n} 之間')
        indices.update(range(a-1,b))
    return sorted(indices)

@app.post('/api/process')
async def api_process(tool:int=Form(...),option:str=Form(''),password:str=Form(''),files:list[UploadFile]=File(...)):
    if tool not in EXT:raise HTTPException(400,'不支援的工具')
    if len(files)>30:raise HTTPException(400,'最多 30 個檔案')
    if tool != 8 and len(files) != 1:raise HTTPException(400,'此工具一次只能處理一個檔案')
    with tempfile.TemporaryDirectory(prefix='dongdong_') as td:
        out=Path(td);paths=[];total=0
        for i,f in enumerate(files):
            ext=Path(f.filename or '').suffix.lower()
            if ext not in EXT[tool]:raise HTTPException(400,'檔案格式不符')
            data=await f.read(MAX+1)
            if len(data)>MAX:raise HTTPException(413,'單一檔案不得超過 40 MB')
            total+=len(data)
            if total>120*1024*1024:raise HTTPException(413,'單次上傳總量不得超過 120 MB')
            p=out/(str(i)+'_'+safe_name(f.filename or 'file'+ext));p.write_bytes(data);paths.append(p)
        try:result=process(tool,paths,out,option,password)
        except (ValueError,fitz.FileDataError,fitz.FileNotFoundError) as e:raise HTTPException(422,str(e))
        except Exception as e:raise HTTPException(500,'處理失敗，請確認來源檔案完整及格式受支援')
        content=result.read_bytes();name=result.name
        from urllib.parse import quote
        return Response(content,media_type='application/zip' if result.suffix=='.zip' else 'application/octet-stream',headers={'Content-Disposition':"attachment; filename*=UTF-8''"+quote(name),'X-Filename':quote(name)})

@app.get('/api/health')
def health():return {'ok':True,'tools':19,'libreoffice':bool(shutil.which('libreoffice'))}
@app.get('/')
def homepage():return FileResponse(ROOT/'index.html',media_type='text/html')
