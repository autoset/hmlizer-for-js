/**
 * �������(hmlizer)
 * ------------------
 * - �Ʒ��ѱ�(.hwp/.hml) ���� ��Ʈ�� ��ũ��Ʈ
 * - hwpctrl.ocx �� �����.
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


// �Ʒ��ѱ� ���α׷� �ʱ�ȭ
hmlizer.core.init();

// HWP ���� ����.
hmlizer.core.open(curPath+"\\sample.hwp");

// ������ �� Ȯ��
WScript.Echo( "�� �������� �� " + hmlizer.core.getPageCount() + "���� �������� �����մϴ�." );

// ù ������ �̹��� ���Ϸ� �����ϱ�
hmlizer.core.createPageImage(curPath+"\\sample_page_1.bmp", 0, 94, 24, "bmp");

// HWP ������ HWPML(HML) ���Ϸ� �����ϱ�
hmlizer.core.saveAsHml(curPath+"\\converted_sample.hml");

// �Ʒ��ѱ� ����
hmlizer.core.quit();
