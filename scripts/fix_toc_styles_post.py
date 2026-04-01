import win32com.client
import os
import re
import sys

def fix_toc_styles_post(file_path):
    print(f"正在对生成的文档进行最终目录样式校正: {file_path}")
    
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        
        constants = win32com.client.constants
        wdTabLeaderDots = getattr(constants, "wdTabLeaderDots", 1)
        wdTabAlignmentRight = getattr(constants, "wdTabAlignmentRight", 2)
        
        doc = word.Documents.Open(os.path.abspath(file_path), ReadOnly=False)
        
        if doc.TablesOfContents.Count == 0:
            print("未检测到自动目录，跳过校正。")
            return
            
        toc = doc.TablesOfContents(1)
        # 获取真实的右对齐边距（默认 A4 约为 415 磅）
        try:
            # 必须从具体节获取，否则如果文档有多节页边距不同，会返回 wdUndefined (很大的负数)
            page_setup = doc.Sections(1).PageSetup
            right_margin_pts = page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin
            if right_margin_pts <= 0 or right_margin_pts > 2000:
                right_margin_pts = 415.5
        except Exception:
            right_margin_pts = 415.5
            
        toc_paragraphs = toc.Range.Paragraphs
        for i in range(1, toc_paragraphs.Count + 1):
            par = toc_paragraphs(i)
            raw_text = str(par.Range.Text).replace("\r", "").replace("\x07", "").replace("\u0007", "")
            if not raw_text.strip():
                continue
                
            # 提取标题文本
            title_part = raw_text
            if "\t" in raw_text:
                title_part = raw_text.rsplit("\t", 1)[0]
            else:
                m = re.search(r'^(.*?)(\d+)\s*$', raw_text.strip())
                if m:
                    title_part = m.group(1)
            title_part = title_part.replace(" ", "").replace("\u3000", "")
            
            # 判断是否是一级标题
            is_level_1 = bool(re.match(r'^第([一二三四五六七八九十百]+|\d+)章', title_part) or title_part in ("参考文献", "致谢", "附录"))
            
            # 1. 设置加粗属性
            par.Range.Font.Bold = is_level_1
            # 强制重写内部所有 Hyperlink 的加粗属性
            try:
                for h_idx in range(1, par.Range.Hyperlinks.Count + 1):
                    par.Range.Hyperlinks(h_idx).Range.Font.Bold = is_level_1
            except Exception:
                pass
                
            # 2. 强制设置制表位（前导点）
            try:
                pf = par.Range.ParagraphFormat
                pf.TabStops.ClearAll()
                pf.TabStops.Add(Position=right_margin_pts, Alignment=wdTabAlignmentRight, Leader=wdTabLeaderDots)
            except Exception:
                pass
                
        doc.Save()
        print("最终目录样式校正完成并保存。")
        
    except Exception as e:
        print(f"校正目录样式时发生错误: {e}")
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except:
                pass
        if word is not None:
            try:
                word.Quit()
            except:
                pass
        
        # 强制清理可能残留的进程
        try:
            import subprocess
            subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
        except Exception:
            pass

if __name__ == "__main__":
    if len(sys.argv) > 1:
        fix_toc_styles_post(sys.argv[1])
    else:
        print("请提供文件路径")