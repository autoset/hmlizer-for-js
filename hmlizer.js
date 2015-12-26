/**
 * 흠라이저(hmlizer)
 * ------------------
 * - 아래한글(.hwp/.hml) 파일 컨트롤 스크립트
 * - hwpctrl.ocx 에 기반함.
 *
 * @author YeonWoong, Jo <caoy@autoset.org>
 * @since 2013
 **/

var fso = WScript.CreateObject("Scripting.FileSystemObject");
var curPath = fso.GetAbsolutePathName(".");

function getModuleInstance(path, varName)
{
	eval((new ActiveXObject("Scripting.FileSystemObject")).OpenTextFile(path, 1).ReadAll());
	return eval(varName);
}

var hmlizer = {
	"core" : getModuleInstance("hmlizer.core.js", "hmlizerCore")
};


// 아래한글 프로그램 초기화
hmlizer.core.init();

// HWP 파일 열기.
hmlizer.core.open(curPath+"\\sample.hwp");

// 페이지 수 확인
WScript.Echo( "이 문서에는 총 " + hmlizer.core.getPageCount() + "개의 페이지가 존재합니다." );

// 첫 페이지 이미지 파일로 생성하기
hmlizer.core.createPageImage(curPath+"\\sample_page_1.bmp", 0, 94, 24, "bmp");

// HWP 파일을 HWPML(HML) 파일로 저장하기
hmlizer.core.saveAsHml(curPath+"\\converted_sample.hml");

// 아래한글 종료
hmlizer.core.quit();
