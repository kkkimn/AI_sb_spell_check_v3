import win32com.client
import pythoncom
import tempfile
import os

def test_extract():
    pythoncom.CoInitialize()
    pptx_path = r"d:\★AI 자동화 시스템 구축 프로그램\테스트4\3. 한기대 스토리보드 맞춤법 & 대본 자동 완성(유료API 사용)_점수 포함 - 복사본\test.pptx"
    
    # Create a dummy PPTX if it doesn't exist
    if not os.path.exists(pptx_path):
        print("Please provide a real PPTX path")
        return
        
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    try:
        presentation = powerpoint.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)
        slide = presentation.Slides(1)
        slide.Export(os.path.join(os.path.dirname(pptx_path), "test_slide.png"), "PNG", 960, 540)
        presentation.Close()
        print("Success!")
    finally:
        powerpoint.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    test_extract()
