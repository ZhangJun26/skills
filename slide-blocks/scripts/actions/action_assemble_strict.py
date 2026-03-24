import win32com.client
import os
import time

def assemble_com():
    base_dir = r"D:\AI\.claude\素材"
    template_path = os.path.join(base_dir, "浅色底模板.pptx")
    hami_path = os.path.join(base_dir, "完整版-哈密市医疗健康数智融合发展规划建议.260202.pptx")
    ai_path = os.path.join(base_dir, "完整版-行业会议-医疗AI的思考与实践-浅色底.251031.pptx")
    out_path = r"D:\AI\.claude\已组装-严格格式版.pptx"
    
    if os.path.exists(out_path):
        try: os.remove(out_path)
        except: pass

    print("启动 PowerPoint 引擎中...")
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True 
    
    try:
        # 打开目标模板
        target_prs = ppt_app.Presentations.Open(template_path)
        
        # 智能寻找包含“标题和正文”占位符的正确版式
        layout = None
        for i in range(1, target_prs.SlideMaster.CustomLayouts.Count + 1):
            lyt = target_prs.SlideMaster.CustomLayouts(i)
            has_title = False
            for j in range(1, lyt.Shapes.Count + 1):
                s = lyt.Shapes(j)
                if s.Type == 14: # msoPlaceholder
                    if s.PlaceholderFormat.Type in [1, 3]: # Title
                        has_title = True
                        break
            if has_title:
                layout = lyt
                # 如果是内容页版式，优先选用
                if "内容" in lyt.Name or "Content" in lyt.Name:
                    break
                    
        if not layout:
            layout = target_prs.SlideMaster.CustomLayouts(2) # 降级方案
            
        def inject(source_slide, index):
            title_text = ""
            shapes_to_copy = []
            
            # 1. 精确识别需要拷贝的形状和标题
            for i in range(1, source_slide.Shapes.Count + 1):
                shp = source_slide.Shapes(i)
                is_bg = False
                
                # 过滤全屏底图
                if shp.Type == 13 and shp.Width > 700 and shp.Height > 400:
                    is_bg = True
                
                is_title = False
                try:
                    # 识别占位符标题
                    if shp.Type == 14 and shp.PlaceholderFormat.Type in [1, 3]:
                        is_title = True
                        if shp.HasTextFrame and shp.TextFrame.HasText:
                            title_text = shp.TextFrame.TextRange.Text
                except Exception:
                    pass
                
                # 识别手工画在顶部的假标题
                if not is_title and shp.HasTextFrame and shp.TextFrame.HasText:
                    if shp.Top < 70 and shp.TextFrame.TextRange.Font.Size >= 18:
                        is_title = True
                        if not title_text:
                            title_text = shp.TextFrame.TextRange.Text
                            
                # 既不是底图也不是标题的内容，才是真正的“积木块”
                if not is_bg and not is_title:
                    shapes_to_copy.append(i)
            
            # 2. 新建一页，强制应用模板版式
            new_slide = target_prs.Slides.AddSlide(target_prs.Slides.Count + 1, layout)
            
            # 3. 注入标题，完全继承模板的标题样式（字体、颜色、位置）
            if title_text and new_slide.Shapes.HasTitle:
                new_slide.Shapes.Title.TextFrame.TextRange.Text = title_text.replace('\v', '').replace('\n', ' ')
                
            # 4. 保留源格式复制积木块，严防色系被篡改
            if shapes_to_copy:
                source_slide.Shapes.Range(shapes_to_copy).Copy()
                
                # 激活模板窗口以执行粘贴
                target_prs.Windows(1).Activate()
                ppt_app.ActiveWindow.View.GotoSlide(new_slide.SlideIndex)
                time.sleep(0.5) # 等待剪贴板响应
                
                # 核心杀招：调用原生 UI 的“保留源格式粘贴 (Keep Source Formatting)”
                ppt_app.CommandBars.ExecuteMso("PasteSourceFormatting")
                
        print("\n>>> 打开素材: 哈密规划")
        hami_prs = ppt_app.Presentations.Open(hami_path, ReadOnly=True)
        for i in range(2, 7):
            print(f"  - 正在严格注入第 {i} 页...")
            inject(hami_prs.Slides(i), i)
        hami_prs.Close()
        
        print("\n>>> 打开素材: 医疗AI实践")
        ai_prs = ppt_app.Presentations.Open(ai_path, ReadOnly=True)
        for i in range(19, 26):
            print(f"  - 正在严格注入第 {i} 页...")
            inject(ai_prs.Slides(i), i)
        ai_prs.Close()
        
        target_prs.SaveAs(out_path)
        target_prs.Close()
        print(f"\n✅ 完美组装！严格保留色系版已保存至:\n{out_path}")
        
    except Exception as e:
        print(f"❌ 发生致命错误: {e}")
    finally:
        ppt_app.Quit()

if __name__ == "__main__":
    assemble_com()
