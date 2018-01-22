// 窗体布局

package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"gitee.com/ying32/govcl/vcl"
	"gitee.com/ying32/govcl/vcl/rtl"
	"gitee.com/ying32/govcl/vcl/types"
	"gitee.com/ying32/govcl/vcl/win"
)

var (
	fSortOrder bool
)

const E_WARN_STR = "生成警告"
const E_ERROT_STR = "生成错误"

type TFormConv struct {
	*vcl.TForm
	icon                    *vcl.TIcon        // ICON
	MainMenu                *vcl.TMainMenu    // 主菜单栏
	FrmAbout                *vcl.TForm        // 关于
	Panel                   *vcl.TPanel       // 布局panel
	Label1, Label2, Label3  *vcl.TLabel       // xlsx路径、输出路径、翻译路径标签
	InputCbox               *vcl.TComboBox    // xlsx路径选择框
	OutOutEdit              *vcl.TEdit        // 输出路径框
	LangEdit                *vcl.TEdit        // 翻译路径框
	Btn1, Btn2              *vcl.TButton      // 按钮(选择路径，生成配置)
	AllChkBox, ChangeChkBox *vcl.TCheckBox    // 全选、选择有变化的
	TestChkBox              *vcl.TCheckBox    // 实验性黑科技
	ListView                *vcl.TListView    // 列表
	PrgBar                  *vcl.TProgressBar // 进度条
	Statusbar               *vcl.TStatusBar   // 底部状态栏
	Pmitem                  *vcl.TPopupMenu   // 右键菜单
	Inifile                 *vcl.TIniFile     // 历史记录
	History                 []string          // 分支路径
}

// 获取程序运行路径
func getCurrentDirectory() string {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		return ""
	}
	return dir
	//return strings.Replace(dir, "\\", "/", -1)
}

/*------------------------private------------------------*/
func (f *TFormConv) getInPutDir() string {
	return f.InputCbox.Text()
}

func (f *TFormConv) getOutPutDir() string {
	return f.OutOutEdit.Text()
}

func (f *TFormConv) getLangDir() string {
	return f.LangEdit.Text()
}

func (f *TFormConv) getParentDir() string {
	inputDir := f.InputCbox.Text()
	parentDir, err := filepath.Abs(inputDir + "\\..")
	if err != nil {
		f.MsgBox("路径错误", "错误")
	}
	return parentDir
}

func (f *TFormConv) MsgBox(text, caption string) {
	if len(text) > 0 {
		vcl.Application.MessageBox(text, caption, win.MB_OK+win.MB_ICONINFORMATION)
	}
}

func (f *TFormConv) updateEdit() {
	dir := f.getParentDir()
	if len(dir) > 0 {
		f.OutOutEdit.SetText(dir + "\\l-xlsx")
		f.LangEdit.SetText(FindLangFolder(dir))
	}
}

func (f *TFormConv) updateProcess() {
	old := f.PrgBar.Position()
	f.PrgBar.SetPosition(old + 1)
}

func (f *TFormConv) loadIni() {
	iniFile := vcl.NewIniFile(`C:\Users\Administrator\Documents\xlsx2lua.ini`)
	f.Inifile = iniFile

	for i := 1; i <= 10; i++ {
		history := iniFile.ReadString("History", fmt.Sprintf("path%d", i), "")
		if len(history) > 0 {
			f.History = append(f.History, history)
		}
	}
}

func (f *TFormConv) saveIni() {
	for i, his := range f.History {
		f.Inifile.WriteString("History", fmt.Sprintf("path%d", i+1), his)
	}
}

// 主菜单
func (f *TFormConv) initFormMenu() {
	mainForm := f.TForm
	mainMenu := vcl.NewMainMenu(f)
	f.MainMenu = mainMenu

	// 不自动生成热键
	mainMenu.SetAutoHotkeys(types.MaManual)
	// 一级菜单
	item := vcl.NewMenuItem(mainForm)
	item.SetCaption("文件(&F)")

	subMenu := vcl.NewMenuItem(mainForm)
	subMenu.SetCaption("新建(&N)")
	subMenu.SetShortCutFromString("Ctrl+N")
	subMenu.SetOnClick(func(vcl.IObject) {
		//fmt.Println("单击了新建")
	})
	item.Add(subMenu)

	subMenu = vcl.NewMenuItem(mainForm)
	subMenu.SetCaption("打开(&O)")
	subMenu.SetShortCutFromString("Ctrl+O")
	item.Add(subMenu)

	subMenu = vcl.NewMenuItem(mainForm)
	subMenu.SetCaption("保存(&S)")
	subMenu.SetShortCutFromString("Ctrl+S")
	item.Add(subMenu)

	// 分割线
	subMenu = vcl.NewMenuItem(mainForm)
	subMenu.SetCaption("-")
	item.Add(subMenu)

	subMenu = vcl.NewMenuItem(mainForm)
	subMenu.SetCaption("退出(&Q)")
	subMenu.SetShortCutFromString("Ctrl+Q")
	subMenu.SetOnClick(func(vcl.IObject) {
		mainForm.Close()
	})
	item.Add(subMenu)

	mainMenu.Items().Add(item)

	item = vcl.NewMenuItem(mainForm)
	item.SetCaption("帮助(&H)")

	subMenu = vcl.NewMenuItem(mainForm)
	subMenu.SetCaption("关于(&A)")
	item.Add(subMenu)
	mainMenu.Items().Add(item)
	subMenu.SetOnClick(func(vcl.IObject) {
		f.FrmAbout.ShowModal()
	})

	// 状态栏
	statusbar := vcl.NewStatusBar(mainForm)
	statusbar.SetParent(mainForm)
	statusbar.SetName("statusbar")
	statusbar.SetSizeGrip(false) // 右下角出现可调整窗口三角形，默认显示
	pnl := statusbar.Panels().Add()
	pnl.SetText("文件数量:0")
	pnl.SetWidth(100)
	pn2 := statusbar.Panels().Add()
	pn2.SetText("有变化的数量:0")
	pn2.SetWidth(200)
	pn3 := statusbar.Panels().Add()
	pn3.SetText("总耗时(ms):0")
	pn3.SetWidth(100)
	f.Statusbar = statusbar
}

func (f *TFormConv) initfrmAbout() {
	frmAbout := vcl.Application.CreateForm()
	frmAbout.ScreenCenter()
	frmAbout.SetCaption("关于")
	frmAbout.SetBorderStyle(types.BsSingle)
	frmAbout.EnabledMaximize(false)
	frmAbout.EnabledMinimize(false)
	frmAbout.SetWidth(405)
	frmAbout.SetHeight(210)
	f.FrmAbout = frmAbout

	about := vcl.NewLabel(frmAbout)
	about.SetParent(frmAbout)
	about.SetAlign(types.AlClient)
	//about.SetTop(frmAbout.ClientHeight() / 2)
	about.SetAutoSize(false)
	about.SetAlignment(types.TaCenter)
	about.SetLayout(types.TlCenter)
	about.SetStyleElements(types.AkRight)
	about.SetCaption("这是一个奇怪的工具\r\ndomi © 2018")

	//	btn := vcl.NewButton(frmAbout)
	//	btn.SetParent(frmAbout)
	//	btn.SetCaption("OK")
	//	btn.SetModalResult(types.MbOK)
	//	btn.SetLeft(frmAbout.ClientWidth() - btn.Width() - 10)
	//	btn.SetTop(frmAbout.ClientHeight() - btn.Height() - 10)
}

// 初始化panel内的布局
func (f *TFormConv) initPanel() {
	mainForm := f.TForm
	pnl := vcl.NewPanel(mainForm)
	pnl.SetParent(mainForm)
	pnl.SetHeight(130)
	pnl.SetAlign(types.AlTop)
	f.Panel = pnl

	_createLabel := func(caption string, left, top int32) *vcl.TLabel {
		label := vcl.NewLabel(mainForm)
		label.SetLeft(left)
		label.SetTop(top)
		label.SetCaption(caption)
		label.SetParent(pnl)
		return label
	}
	_createEdit := func(caption string, left, top int32) *vcl.TEdit {
		edit := vcl.NewEdit(mainForm)
		edit.SetLeft(left)
		edit.SetTop(top)
		edit.SetText(caption)
		edit.SetWidth(300)
		edit.SetReadOnly(true)
		edit.SetParent(pnl)
		return edit
	}
	_createBtn := func(caption string, left, top int32) *vcl.TButton {
		btn := vcl.NewButton(mainForm)
		btn.SetParent(pnl)
		btn.SetLeft(left)
		btn.SetTop(top)
		btn.SetWidth(100)
		btn.SetCaption(caption)
		return btn
	}
	_createChkBox := func(caption string, left, top int32) *vcl.TCheckBox {
		chkBox := vcl.NewCheckBox(mainForm)
		chkBox.SetParent(pnl)
		chkBox.SetLeft(left)
		chkBox.SetTop(top)
		chkBox.SetCaption(caption)
		return chkBox
	}

	// 第1行
	{
		left, top := int32(10), int32(20)
		f.Label1 = _createLabel("配置路径：", left, top)

		// 路径input
		left += f.Label1.Width() + 5
		cbox := vcl.NewComboBox(mainForm)
		cbox.SetParent(mainForm)
		cbox.SetLeft(left)
		cbox.SetTop(top)
		cbox.SetWidth(300)
		cbox.SetStyle(types.CsOwnerDrawFixed)
		for _, his := range f.History {
			cbox.Items().Add(his)
		}
		cbox.SetItemIndex(0)
		f.InputCbox = cbox

		// btn
		top -= 5
		left += cbox.Width() + 10
		f.Btn1 = _createBtn("选择路径", left, top)

		left += f.Btn1.Width() + 10
		f.Btn2 = _createBtn("生成配置", left, top)
	}

	// 第2行
	{
		left, top := int32(10), f.Label1.Top()+f.Label1.Height()+20
		f.Label2 = _createLabel("输出路径：", left, top)
		left += f.Label2.Width() + 5
		f.OutOutEdit = _createEdit("", left, top)
	}

	// 第3行
	{
		left, top := int32(10), f.Label2.Top()+f.Label2.Height()+20
		f.Label3 = _createLabel("翻译路径：", left, top)
		left += f.Label3.Width() + 5
		f.LangEdit = _createEdit("", left, top)
		left += f.LangEdit.Width() + 5

		prgLable := _createLabel("生成进度：", left, top)
		left += prgLable.Width() + 5
		prgbar := vcl.NewProgressBar(mainForm)
		prgbar.SetParent(mainForm)
		prgbar.SetBounds(left, top, 200, 20)
		prgbar.SetMin(0)
		prgbar.SetPosition(0)
		prgbar.SetOrientation(types.PbHorizontal)
		f.PrgBar = prgbar
	}

	// 第4行
	{
		left, top := int32(10), f.Label3.Top()+f.Label3.Height()+10
		f.AllChkBox = _createChkBox("全选", left, top)
		left = left + f.AllChkBox.Width() + 10
		f.ChangeChkBox = _createChkBox("选择有变化的", left, top)
		left = left + f.ChangeChkBox.Width() + 10
		f.TestChkBox = _createChkBox("实验性特性", left, top)
	}
	f.updateEdit()
}

func (f *TFormConv) initListView() {
	// TPopupMenu
	mainForm := f.TForm
	pm := vcl.NewPopupMenu(mainForm)
	pmitem := vcl.NewMenuItem(mainForm)
	pmitem.SetCaption("打开文件")
	pm.Items().Add(pmitem)

	pmitem1 := vcl.NewMenuItem(mainForm)
	pmitem1.SetCaption("打开文件所在目录")
	pm.Items().Add(pmitem1)

	pmitem2 := vcl.NewMenuItem(mainForm)
	pmitem2.SetCaption("打开输出目录")
	pm.Items().Add(pmitem2)

	pmitem3 := vcl.NewMenuItem(mainForm)
	pmitem3.SetCaption("显示错误")
	pm.Items().Add(pmitem3)
	f.Pmitem = pm

	// 生成结果列表
	imgList := vcl.NewImageList(mainForm)
	//imgList.SetHeight(100)
	imgList.SetWidth(1)
	lv1 := vcl.NewListView(mainForm)
	lv1.SetParent(mainForm)
	lv1.SetWidth(mainForm.ClientWidth())
	lv1.SetAlign(types.AlClient)
	//lv1.SetClientWidth(300)
	lv1.SetSmallImages(imgList)
	lv1.SetRowSelect(true)
	lv1.SetReadOnly(true)
	lv1.SetGridLines(true)
	lv1.SetViewStyle(types.VsReport)
	lv1.Font().SetName("微软雅黑")
	lv1.Font().SetSize(10)
	lv1.SetCheckboxes(true)
	lv1.SetPopupMenu(pm)
	f.ListView = lv1

	addCol := func(caption string, width int32, autosize bool) {
		col := lv1.Columns().Add()
		col.SetCaption(caption)
		col.SetWidth(width)
		col.SetAutoSize(autosize)
		if !autosize {
			col.SetMaxWidth(width)
			col.SetMinWidth(width)
		}
		col.SetAlignment(types.TaLeftJustify)
	}
	addCol("文件名", lv1.ClientWidth()-317, true)
	addCol("文件状态", 100, false)
	addCol("生成结果", 200, false)

	// 右键菜单相应
	pmitem.SetOnClick(func(vcl.IObject) {
		item := f.ListView.Selected()
		rtl.SysOpen(item.Caption())
		//cmdStr := exec.Command("cmd", "/C start "+item.Caption())
		//go cmdStr.Run()
	})
	pmitem1.SetOnClick(func(vcl.IObject) {
		item := f.ListView.Selected()
		rtl.SysOpen(rtl.ExtractFilePath(item.Caption()))
	})
	pmitem2.SetOnClick(func(vcl.IObject) {
		item := f.ListView.Selected()
		idx := int(item.Data())
		dir := f.getOutPutDir() + "\\" + Convs[idx].FolderName
		println(dir)
		rtl.SysOpen(dir)
	})
	pmitem3.SetOnClick(func(vcl.IObject) {
		item := f.ListView.Selected()
		idx := int(item.Data())
		f.MsgBox(Convs[idx].formatErr(), "生成结果")
	})
}

// 设置控件的事件
func (f *TFormConv) setEvent() {
	cbox := f.InputCbox
	lv1 := f.ListView
	btn1, btn2 := f.Btn1, f.Btn2
	allChkBox := f.AllChkBox
	cbox.SetOnChange(func(vcl.IObject) {
		if cbox.ItemIndex() != -1 {
			f.updateEdit()
			f.LoadXlxs()
		}
	})

	// listview 排序
	lv1.SetOnCompare(lvTraiCompare)
	lv1.SetOnColumnClick(func(sender vcl.IObject, column *vcl.TListColumn) {
		// 按柱头索引排序, lcl兼容版第二个参数永远为 column
		fSortOrder = !fSortOrder
		lv1.CustomSort(0, int(column.Index()))
	})
	lv1.SetOnDblClick(func(sender vcl.IObject) {
		item := f.ListView.Selected()
		item.SetChecked(!item.Checked())
	})
	lv1.SetOnAdvancedCustomDrawItem(func(sender *vcl.TListView, item *vcl.TListItem, state types.TCustomDrawState, Stage types.TCustomDrawStage, defaultDraw *bool) {
		canvas := sender.Canvas()
		font := canvas.Font()
		i := int(item.Index())
		if i%2 == 0 {
			canvas.Brush().SetColor(0x02F0EEF7)
		}

		resStr := item.SubItems().Strings(1)
		if resStr == E_ERROT_STR {
			canvas.Brush().SetColor(types.ClRed)
			font.SetColor(types.ClSilver)
		} else if resStr == E_WARN_STR {
			canvas.Brush().SetColor(types.ClYellow)
			font.SetColor(types.ClSilver)
		}
	})

	// button
	btn1.SetOnClick(func(vcl.IObject) {
		options := types.TSelectDirExtOpts(rtl.Include(0, types.SdNewFolder, types.SdShowEdit, types.SdNewUI))
		if ok, dir := vcl.SelectDirectory2("选择配置路径", "C:/", options, nil); ok {
			f.History = append(f.History, dir)
			cbox.SetText(dir)
			idx := cbox.Items().Add(dir)
			cbox.SetItemIndex(idx)

			f.updateEdit()
			f.LoadXlxs()
		}
	})
	btn2.SetOnClick(func(vcl.IObject) {
		count := lv1.Items().Count()
		idxs := make(map[int]bool, count)
		var i int32
		for i = 0; i < count; i++ {
			item := lv1.Items().Item(i)
			if item.Checked() {
				idxs[int(item.Data())] = true
			}
		}
		if len(idxs) > 0 {
			btn1.SetEnabled(false)
			btn2.SetEnabled(false)
			cbox.SetEnabled(false)

			go startConv(idxs)
		} else {
			f.MsgBox("请选择配置", "通知")
		}
	})

	allChkBox.SetOnClick(func(vcl.IObject) {
		var i int32
		for i = 0; i < lv1.Items().Count(); i++ {
			lv1.Items().Item(i).SetChecked(allChkBox.Checked())
		}
		if allChkBox.Checked() {
			f.PrgBar.SetMax(int32(len(Convs)))
		}
	})

	f.ChangeChkBox.SetOnClick(func(vcl.IObject) {
		listView := f.ListView
		var i int32
		count := listView.Items().Count()
		for i = 0; i < count; i++ {
			item := listView.Items().Item(i)
			idx := int(item.Data())
			if Convs[idx].hasChanged() {
				item.SetChecked(f.ChangeChkBox.Checked())
			}
		}
	})
}

// 排序
func lvTraiCompare(sender vcl.IObject, item1, item2 *vcl.TListItem, data int32, compare *int32) {
	var s1, s2 string
	if data != 0 {
		s1 = item1.SubItems().Strings(data - 1)
		s2 = item2.SubItems().Strings(data - 1)
	} else {
		s1 = item1.Caption()
		s2 = item2.Caption()
	}
	if fSortOrder {
		*compare = int32(strings.Compare(s1, s2))
	} else {
		*compare = -int32(strings.Compare(s1, s2))
	}
}

/*------------------------public------------------------*/
func CreateMainForm() *TFormConv {
	form := new(TFormConv)

	// icon
	icon := vcl.NewIcon()
	icon.LoadFromResourceID(rtl.MainInstance(), 3)
	vcl.Application.Initialize()
	vcl.Application.SetMainFormOnTaskBar(false)
	vcl.Application.SetIcon(icon)

	mainForm := vcl.Application.CreateForm()
	mainForm.SetCaption("xlsx2lua")
	mainForm.ScreenCenter()
	mainForm.SetPosition(types.PoScreenCenter)
	mainForm.EnabledMaximize(false)
	//mainForm.SetBorderStyle(types.BsSingle)
	mainForm.SetWidth(1024)
	mainForm.SetHeight(800)
	mainForm.SetDoubleBuffered(true)

	form.icon = icon
	form.TForm = mainForm
	form.History = make([]string, 0, 10)
	return form
}

// 创建窗体内的控件
func (f *TFormConv) CreateControl() {
	f.loadIni()
	f.initFormMenu()
	f.initfrmAbout()
	f.initPanel()
	f.initListView()
	f.setEvent()
}

func (f *TFormConv) LoadXlxs() {
	dir := f.InputCbox.Text()
	if len(dir) > 0 {
		err := WalkXlsx(dir)
		if err == nil {
			listView := f.ListView
			listView.Items().Clear()
			listView.Items().BeginUpdate()

			convsLen := int32(len(Convs))
			changeCount := 0
			for i, conv := range Convs {
				item := listView.Items().Add()
				// 第一列为Caption属性所管理
				isChange := conv.hasChanged()
				item.SetChecked(isChange)
				item.SetCaption(conv.AbsPath)
				if isChange {
					changeCount += 1
					item.SubItems().Add("配置有变化")
				} else {
					item.SubItems().Add("-")
				}
				item.SubItems().Add("-")
				item.SetData(uintptr(i))
			}
			listView.Items().EndUpdate()

			listView.CustomSort(0, int(1)) // 按是否变化排序列表
			f.ChangeChkBox.SetChecked(true)

			f.PrgBar.SetMax(int32(changeCount))
			f.PrgBar.SetPosition(0)

			f.Statusbar.Panels().Items(0).SetText(fmt.Sprintf("文件数量：%d", convsLen))
			f.Statusbar.Panels().Items(1).SetText(fmt.Sprintf("有变化的数量：%d", changeCount))
			f.saveIni()
		} else {
			f.MsgBox(err.Error(), "加载配置错误")
		}
	}
}

func (f *TFormConv) ConvResult(idxs map[int]bool, startTime time.Time) {
	f.Btn2.SetEnabled(true)
	f.Btn1.SetEnabled(true)
	f.InputCbox.SetEnabled(true)
	listView := f.ListView

	var i int32
	var c *XlsxConv
	var errCount, warnCount int
	count := listView.Items().Count()
	for i = 0; i < count; i++ {
		item := listView.Items().Item(i)
		idx := int(item.Data())
		if _, ok := idxs[idx]; ok {
			c = Convs[idx]
			//item.SubItems().BeginUpdate()
			if c.hasError(E_ERROR) {
				errCount++
				item.SubItems().SetStrings(1, E_ERROT_STR)
			} else if c.hasError(E_WARN) {
				warnCount++
				item.SubItems().SetStrings(0, "-")
				item.SubItems().SetStrings(1, E_WARN_STR)
			} else {
				item.SubItems().SetStrings(0, "-")
				item.SubItems().SetStrings(1, fmt.Sprintf("耗时(ms):%d", c.Msec))
			}
			//item.SubItems().EndUpdate()
		}
	}
	f.PrgBar.SetPosition(0)
	f.Statusbar.Panels().Items(2).SetText(fmt.Sprintf("总耗时(ms)：%d", int(time.Now().Sub(startTime).Nanoseconds()/1e6)))

	f.MsgBox(fmt.Sprintf("错误：%d条，警告：%d条", errCount, warnCount), "生成结果")
}
