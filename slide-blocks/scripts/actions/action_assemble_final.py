import win32com.client
import os
import time

def assemble_com():
    base_dir = r"D:\AI\.claude\素材"
    template_path = os.path.join(base_dir, "浅色底模板.pptx")
    hami_path = os.path.join(base_dir, "完整版-哈密市医疗健康数智融合发展规划建议.260202.pptx")
    ai_path = os.path.join(base_dir, "完整版-行业会议-医疗AI的思考与实践-浅色底.251031.pptx")
    out_path = r"D:\AI\.claude\已组装-严格格式版_最终修复.pptx"
    
    if os.path.exists(out_path):
        try: os.remove(out_path)
        except: pass

    print("启动 PowerPoint 引擎中...")
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True 
    
    try:
        target_prs = ppt_app.Presentations.Open(template_path)
        
        # 核心修复 1：把带有 "XXX" 的这页当做真正的模板页
        # 因为它不是母版，而是实打实的一页幻灯片。
        template_slide = target_prs.Slides(2) 
        
        def inject(source_slide, index):
            title_text = ""
            shapes_to_copy = []
            
            # 从源文件提取文字和要复制的形状
            for i in range(1, source_slide.Shapes.Count + 1):
                shp = source_slide.Shapes(i)
                is_bg = (shp.Type == 13 and shp.Width > 700 and shp.Height > 400)
                is_title = False
                
                try:
                    if shp.Type == 14 and shp.PlaceholderFormat.Type in [1, 3]:
                        is_title = True
                        if shp.HasTextFrame and shp.TextFrame.HasText:
                            title_text = shp.TextFrame.TextRange.Text
                except Exception: pass
                
                if not is_title and shp.HasTextFrame and shp.TextFrame.HasText:
                    if shp.Top < 70 and shp.TextFrame.TextRange.Font.Size >= 18:
                        is_title = True
                        if not title_text: title_text = shp.TextFrame.TextRange.Text
                            
                if not is_bg and not is_title:
                    shapes_to_copy.append(i)
            
            # 核心修复 2：不是 AddSlide，而是 Duplicate 那页 "XXX"
            new_slide_range = template_slide.Duplicate()
            new_slide = new_slide_range.Item(1)
            # 把它移动到最后
            new_slide.MoveTo(target_prs.Slides.Count)
            
            # 核心修复 3：寻找 "XXX" 文本框，替换为真实的源标题，从而 100% 继承 "XXX" 的样式
            if title_text:
                for j in range(1, new_slide.Shapes.Count + 1):
                    ts = new_slide.Shapes(j)
                    if ts.HasTextFrame and ts.TextFrame.HasText:
                        if "XXX" in ts.TextFrame.TextRange.Text:
                            # 替换文本，保留源样式
                            ts.TextFrame.TextRange.Text = title_text.replace('\v', '').replace('\n', ' ')
                            break
            
            # 粘贴内容并保持源格式（色系）
            if shapes_to_copy:
                source_slide.Shapes.Range(shapes_to_copy).Copy()
                target_prs.Windows(1).Activate()
                ppt_app.ActiveWindow.View.GotoSlide(new_slide.SlideIndex)
                time.sleep(0.5) 
                ppt_app.CommandBars.ExecuteMso("PasteSourceFormatting")
                
        print("\n>>> 打开素材: 哈密规划")
        hami_prs = ppt_app.Presentations.Open(hami_path, ReadOnly=True)
        for i in range(2, 7):
            print(f"  - 正在克隆第 {i} 页...")
            inject(hami_prs.Slides(i), i)
        hami_prs.Close()
        
        print("\n>>> 打开素材: 医疗AI实践")
        ai_prs = ppt_app.Presentations.Open(ai_path, ReadOnly=True)
        for i in range(19, 26):
            print(f"  - 正在克隆第 {i} 页...")
            inject(ai_prs.Slides(i), i)
        ai_prs.Close()
        
        # 删除原本用作模板的第 1-4 页，只留下新生成的组装页
        for i in range(4, 0, -1):
            target_prs.Slides(i).Delete()
            
        target_prs.SaveAs(out_path)
        target_prs.Close()
        print(f"\n✅ 完美组装！XXX 标题继承版已保存至:\n{out_path}")
        
    except Exception as e:
        print(f"❌ 发生致命错误: {e}")
    finally:
        ppt_app.Quit()

if __name__ == "__main__":
    assemble_com()
