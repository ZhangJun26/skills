import win32com.client
import os
import time

def assemble_com():
    base_dir = r"D:\AI\.claude\素材"
    template_path = os.path.join(base_dir, "浅色底模板.pptx")
    hami_path = os.path.join(base_dir, "完整版-哈密市医疗健康数智融合发展规划建议.260202.pptx")
    ai_path = os.path.join(base_dir, "完整版-行业会议-医疗AI的思考与实践-浅色底.251031.pptx")
    out_path = r"D:\AI\.claude\已组装-原生无损版.pptx"
    
    # 清理旧文件
    if os.path.exists(out_path):
        try: os.remove(out_path)
        except: pass

    print("启动 PowerPoint 引擎中 (可能需要几秒钟)...")
    # 启动 PowerPoint 应用
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    # 让 PPT 窗口可见以确保剪贴板操作不出错
    ppt_app.Visible = True 
    
    try:
        # 打开模板
        target_prs = ppt_app.Presentations.Open(template_path)
        # 获取模板的版式 (2 通常是 标题+正文 版式)
        layout = target_prs.SlideMaster.CustomLayouts(2) 
        
        def inject(source_slide, index):
            print(f"  - 正在剥离并注入源文件第 {index} 页...")
            title_text = ""
            shapes_to_copy = []
            
            # 1. 精准识别标题与背景
            for i in range(1, source_slide.Shapes.Count + 1):
                shp = source_slide.Shapes(i)
                # 过滤可能的全屏底图 (Type 13 是图片, 宽度 > 700 磅通常是背景)
                is_bg = (shp.Type == 13 and shp.Width > 700 and shp.Height > 400)
                
                is_title = False
                if shp.HasTextFrame and shp.TextFrame.HasText:
                    # 识别占位符标题
                    if shp.Type == 14 and ("Title" in shp.Name or "标题" in shp.Name):
                        is_title = True
                        title_text = shp.TextFrame.TextRange.Text
                    # 识别手工画在顶部的假标题
                    elif not title_text and shp.Top < 60 and shp.TextFrame.TextRange.Text.strip(): 
                        is_title = True
                        title_text = shp.TextFrame.TextRange.Text
                        
                if not is_bg and not is_title:
                    shapes_to_copy.append(i)
            
            # 2. 在模板新建一页
            new_slide = target_prs.Slides.AddSlide(target_prs.Slides.Count + 1, layout)
            
            # 3. 注入模板标准的标题
            if title_text and new_slide.Shapes.HasTitle:
                new_slide.Shapes.Title.TextFrame.TextRange.Text = title_text.replace('\v', '').replace('\n', ' ')
                
            # 4. 原生复制粘贴积木块
            if shapes_to_copy:
                source_slide.Shapes.Range(shapes_to_copy).Copy()
                # 使用原生 Paste 自动处理颜色和资源关联，绝不损坏文件
                new_slide.Shapes.Paste()
                
        # --- 处理文件 1 ---
        print("\n>>> 打开素材: 哈密规划")
        hami_prs = ppt_app.Presentations.Open(hami_path, ReadOnly=True)
        for i in range(2, 7): # 第2-6页
            inject(hami_prs.Slides(i), i)
        hami_prs.Close()
        
        # --- 处理文件 2 ---
        print("\n>>> 打开素材: 医疗AI实践")
        ai_prs = ppt_app.Presentations.Open(ai_path, ReadOnly=True)
        for i in range(19, 26): # 第19-25页
            inject(ai_prs.Slides(i), i)
        ai_prs.Close()
        
        # 保存并关闭
        target_prs.SaveAs(out_path)
        target_prs.Close()
        print(f"\n✅ 完美组装！原生无损文件已保存至:\n{out_path}")
        
    except Exception as e:
        print(f"❌ 发生致命错误: {e}")
    finally:
        ppt_app.Quit()

if __name__ == "__main__":
    assemble_com()
