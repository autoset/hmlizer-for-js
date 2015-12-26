/**
 * 흠라이저(hmlizer)
 * ------------------
 * - 아래한글(.hwp/.hml) 파일 컨트롤 스크립트
 * - hwpctrl.ocx 에 기반함.
 *
 * @author YeonWoong, Jo <caoy@autoset.org>
 * @since 2013
 **/
var hmlizerCore = new (function()
{
	var App = null;
	var Docs = null;
	var hwpCtrl = null;

	this.init = function()
	{
		try
		{
			App = WScript.CreateObject("HWPFrame.HwpObject.1");
			Docs = App.XHwpDocuments;
			hwpCtrl = Docs.Application;

			hwpCtrl.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample");
		}
		catch(ex)
		{
			WScript.Echo(ex.message);
		}
	}

	this.open = function(path, arg1, arg2)
	{
		if (!hwpCtrl)
			return ;

		try
		{
			hwpCtrl.Open(path,arg1||"",arg2||"");
		}
		catch(ex)
		{
		}
	}

	this.getPageCount = function()
	{
		if (!hwpCtrl)
			return false;

		return hwpCtrl.PageCount;
	}

	this.createPageImage = function(bmpPath, arg1, arg2, arg3, arg4)
	{
		if (!hwpCtrl)
			return ;

		try
		{
			hwpCtrl.CreatePageImage(bmpPath, arg1||0, arg2||300, arg3||24, arg4||"bmp");	
		}
		catch (ex)
		{
		}
	}

	this.saveAsHml = function(path, bEqOtmz)
	{
		if (!hwpCtrl)
			return ;

		if (bEqOtmz == undefined)
			bEqOtmz = false;

		hwpCtrl.Open(path,"","");

		if (bEqOtmz)
		{
			// 수식 바로잡기(수식 크기가 비정상적으로 작게 보이는 현상을 바로잡습니다)
			try
			{
				var headCtrl = hwpCtrl.HeadCtrl;
				while (true)
				{
					if (headCtrl == null)
					{
						break;
					}
					if (headCtrl.CtrlID == "eqed")
					{
						var anchorSet = headCtrl.GetAnchorPos(0);
						var list = anchorSet.Item("List");
						hwpCtrl.SetPos(list, anchorSet.Item("Para"), anchorSet.Item("Pos"));
						hwpCtrl.FindCtrl();
						var  action = hwpCtrl.CreateAction("EquationPropertyDialog");
						var  eqSet = action.CreateSet();
						action.GetDefault(eqSet);
						eqSet.SetItem("Version", "Equation Version 60");
						action.Execute(eqSet);
						hwpCtrl.UnSelectCtrl();
						hwpCtrl.Run("MoveRight");
					}
					headCtrl = headCtrl.Next;
				}
			}
			catch (ex)
			{
			}

		}

		var savedPath = path.replace(/\.hwp$/g,".hml");


		try
		{
			hwpCtrl.SaveAs(savedPath, "HWPML2X", "");
			hwpCtrl.Clear(1);
		}
		catch(ex)
		{
		}

		return savedPath;
	}

	this.saveAs = function(path, destPath, fileFormat)
	{
		if (!hwpCtrl)
			return ;

		try
		{
			hwpCtrl.Open(path,"","");
			hwpCtrl.SaveAs(destPath, fileFormat || "HWPML2X", "");
			hwpCtrl.Clear(1);
		}
		catch(ex)
		{
		}

		return destPath;
	}

	this.clear = function()
	{
		if (!hwpCtrl)
			return ;

		try
		{
			hwpCtrl.Clear(1);
		}
		catch(ex)
		{
		}
	}

	this.quit = function()
	{
		if (!hwpCtrl)
			return ;

		try
		{
			hwpCtrl.Clear(1);
			hwpCtrl.Run("FileQuit");
		}
		catch(ex)
		{
		}
	}
});
