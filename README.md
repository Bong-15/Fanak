# Fanak 변환기 (Fanak) 개발 문서

탐색기에서 파일을 끌다가 키 하나를 누르면 원형 메뉴가 열리고, 이미지·PDF를 변환하거나 이미지를 자르고 DPI를 바꾸거나 문서를 Markdown으로 바꾸는 Windows 앱입니다.

**[최신 버전 내려받기](https://github.com/Bong-15/Fanak/releases/latest)** zip을 풀어 `Fanak.exe`를 실행하면 됩니다. 사용법은 그 안의 `사용설명서.txt`에 있습니다.

현재 버전은 **261103 ver 1.0**입니다. 표기는 `src/SettingsWindow.cs`의 `Version`과 `src/AssemblyInfo.cs`, `package/사용설명서.txt`에 있습니다.

## 빌드

Windows에 포함된 C# 컴파일러를 쓰므로 따로 설치할 것이 없습니다. 결과는 `dist/Fanak`에 만들어지며, 이 폴더만 배포하면 됩니다.

```bash
powershell -ExecutionPolicy Bypass -File build.ps1
```

## 코드 구성

- `src/MouseWatcher.cs`: 전역 마우스 훅으로 호출 키를 누른 드래그를 감지합니다.
- `src/RadialWindow.cs`: 원형 메뉴이자 드롭 대상입니다.
- `src/EditWindows.cs`: 자르기 창, DPI 창, 창 공통 도구입니다.
- `src/StatusWindow.cs`: 화면 구석의 진행·완료·실패 카드입니다.
- `src/SettingsWindow.cs`: 설정 창(일반, 사용법, 정보 탭)입니다.
- `src/Settings.cs`: 설정 저장, 자동 실행, 호출 키 이름을 맡습니다.
- `src/Theme.cs`: 모든 창이 함께 쓰는 스타일입니다.
- `src/Converter.cs`: 지원 형식 판단, 변환, DPI 확인, JPG DPI 수정을 맡습니다. 이미지는 WIC, PDF 읽기는 Windows.Data.Pdf를 쓰고, PDF 쓰기는 직접 구현했습니다.
- `src/Markdown.cs`: MD 변환의 입구와 공통 출력(서식, 표, 각주, 그림 폴더)입니다.
- `src/OfficeReaders.cs`: `.docx`, `.pptx`, `.xlsx`, `.hwpx`를 읽습니다. 모두 XML을 담은 zip이라 해당 프로그램이 필요 없습니다.
- `src/HwpReader.cs`: 나머지 한글 형식을 읽습니다. 확장자가 아니라 파일 첫 바이트로 HWP 5.0(`.hwp`, `.hwt`), HWP 3.x, HWPML(`.hml`, `.hwpml`), zip(`.hwpx`, `.hwtx`, `.owpml`)을 가립니다.
- `src/Hwp3Codes.cs`: HWP 3.x의 특수 기호 코드표입니다. 한글로 KS X 1001 전체를 저장한 파일에서 읽어 낸 값입니다.
- `src/TextDocs.cs`: `.txt`, `.csv`, `.tsv`, `.html`, `.rtf`를 읽고, CSV를 `.xlsx`로 씁니다. 인코딩(UTF-8/CP949)을 판별하고, Word에 넘길 HTML에서 외부 연결을 지웁니다.
- `src/OfficePdf.cs`: 설치된 Word, PowerPoint, Excel, 한글을 창 없이 띄워 PDF로 내보내고, 한글로 DOCX·HWP 변환을 합니다. 앱에서 유일하게 Windows 외의 프로그램이 필요한 기능이라, 없는 PC에서는 메뉴에 나오지 않습니다.
- `src/PdfText.cs`: PDF 구조를 직접 읽어 글자, 선, 그림의 위치를 뽑습니다. Windows.Data.Pdf는 페이지를 그리기만 하고 글자를 주지 않기 때문입니다.
- `src/PdfToMd.cs`: 그 위치로 줄, 단, 표, 문단, 제목을 복원합니다. 글자를 읽을 수 없는 쪽은 Windows.Media.Ocr로 읽습니다.
- `src/Program.cs`: 진입점, 트레이 아이콘, 변환 대기열, 명령줄 모드가 있습니다.
- `src/AssemblyInfo.cs`: 파일 속성에 보이는 제품 이름과 버전입니다.
- `tools/MakeIcon.ps1`: 사막여우 아이콘을 그려 `src/Fanak.ico`를 만듭니다. 그림을 바꿀 때만 실행합니다.
- `package/`: 빌드할 때 배포 폴더로 복사되는 파일입니다.

## 저장 위치

- **설정:** `Fanak.exe` 옆의 `settings.ini`. 폴더에 쓸 수 없으면 `%LOCALAPPDATA%\Fanak\settings.ini`.
- **변환 기록:** `%LOCALAPPDATA%\Fanak\fanak.log`. 파일 경로가 담기므로 일부러 배포 폴더 밖에 둡니다. 1MB를 넘으면 새로 시작합니다.
- **자동 실행:** `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`의 `FanakConvert`. 켰을 때만 만들어집니다.

## DPI 처리

- 저장한 뒤 파일을 다시 읽어 값을 비교합니다. 다르면 결과를 지우고 실패로 알립니다.
- JPG를 JPG로 저장하면서 DPI만 바꿀 때는 다시 압축하지 않고 JFIF와 EXIF의 해상도 값만 수정합니다.
- 다른 경우에는 새 DPI로 인코딩합니다. PNG, TIFF, BMP는 무손실이라 픽셀이 그대로입니다.
- 포토샵 전용 해상도 정보(APP13)와 XMP는 수정하지 않습니다. 포토샵에서 만든 JPG를 포토샵에서 열면 예전 값이 보일 수 있습니다.

## 명령줄 모드

```bash
Fanak.exe --convert PNG --dpi 300 --resample --crop 0,0,100,100 파일
```

`--dpi`, `--resample`, `--crop`은 선택 사항이고, 형식 자리에 `ORIGINAL`을 쓰면 원본 형식을 유지합니다. 문서를 Markdown으로 바꿀 때는 형식 자리에 `MD`를 씁니다. 실행 중인 앱은 `Fanak.exe --quit`로 끌 수 있습니다.

## 알려진 제한

사용자용 제한 사항은 [package/사용설명서.txt](package/사용설명서.txt)에 있습니다. 코드 서명이 없어서 인터넷으로 받은 경우 Windows의 PC 보호 경고가 뜰 수 있습니다.
