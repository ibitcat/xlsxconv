// 一次xlsx转换

package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"
	"unicode"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// error level
const (
	E_NONE   = iota
	E_NOTICE //通知
	E_WARN   //警告
	E_ERROR  //错误
)

type ErrorInfo struct {
	Level  int
	ErrMsg string
}

type FieldInfo struct {
	Name string // 字段名
	Type string // 字段类型
	Mode string // 生成方式(s=server,c=client,d=双端,r=策划)
}

type LangSheet struct {
	SheetRows [][]string     // 翻译文件xlsx数据
	FieldRef  map[string]int // 翻译字段反索引
	IdRef     map[string]int // 翻译id反索引
}

// 一个xlsx转换
type XlsxConv struct {
	AbsPath     string            // 文件绝对路径
	RelPath     string            // 文件相对路径(例如task\task.xlsx)
	FolderName  string            // 文件夹路径(例如task\)
	FileName    string            // 文件名（例如：tast.xlsx）
	ModTime     uint64            // 最后修改时间（毫秒）
	RecordTime  uint64            // 记录的时间
	Fields      map[int]FieldInfo // 字段信息
	hasSrvField bool              // 是否有服务器字段
	Lang        *LangSheet        // 翻译
	Msec        int               // 耗时（毫秒）
	Errs        []ErrorInfo       // 错误信息
	checkOnly   bool              // 仅仅检查配置错误，并不生成
}

func checkJson(text string) error {
	var temp interface{}
	err := json.Unmarshal([]byte(text), &temp)
	if err == nil {
		switch temp.(type) {
		case map[string]interface{}:
		case []interface{}:
		default:
			err = errors.New("json格式错误")
		}
	}
	return err
}

// 实验性特性
func checkAscii(srcStr, desStr string) bool {
	srcBytes := make([]byte, 0, len(srcStr))
	for _, r := range srcStr {
		if r > 0x20 && r <= 0x7f && r != 0x2C && r != 0x2e { //忽略中英文逗号、句号的区别
			srcBytes = append(srcBytes, byte(r))
		}
	}
	byteLen := len(srcBytes)
	if byteLen > 0 {
		var idx int = 0
		for _, r := range desStr {
			if r > 0x20 && r <= 0x7f && byte(r) == srcBytes[idx] {
				idx++
				if idx >= byteLen {
					break
				}
			}
		}
		if byteLen != idx {
			return false
		}
	}
	return true
}

func isChineseChar(str string) bool {
	for _, r := range str {
		if unicode.Is(unicode.Scripts["Han"], r) {
			return true
		}
	}
	return false
}

func (c *XlsxConv) hasError(lv int) bool {
	if c.Errs == nil {
		return false
	}
	for _, e := range c.Errs {
		if e.Level >= lv {
			return true
		}
	}
	return false
}

func (c *XlsxConv) formatErr() string {
	var str string
	if c.Errs != nil {
		warnStr := make([]string, 0, len(c.Errs))
		errStr := make([]string, 0, len(c.Errs))
		for _, e := range c.Errs {
			if e.Level == E_WARN {
				warnStr = append(warnStr, e.ErrMsg)
			} else if e.Level == E_ERROR {
				errStr = append(errStr, e.ErrMsg)
			}
		}

		if len(warnStr) > 0 {
			str += fmt.Sprintf("[警告%d条]：\r\n", len(warnStr)) + strings.Join(warnStr, "\r\n")
		}
		if len(errStr) > 0 {
			str += fmt.Sprintf("[错误%d条]：\r\n", len(errStr)) + strings.Join(errStr, "\r\n")
		}
	}
	return str
}

func (c *XlsxConv) hasChanged() bool {
	return c.ModTime != c.RecordTime
}

// 读取翻译xlsx文件
func (c *XlsxConv) loadLangXlsx() {
	langDir := mainForm.getLangDir()
	if len(langDir) == 0 {
		return
	}
	fileName := strings.Replace("\\"+c.RelPath, "\\", "$", -1) // \替换成$
	xlFile, err := excelize.OpenFile(langDir + "\\" + fileName)
	if err != nil {
		return
	}

	//sheetName := strings.TrimSuffix(p, ".xlsx")
	sheetName := xlFile.GetSheetName(1)
	sheet := xlFile.GetRows(sheetName)
	if len(sheet) == 0 {
		c.Errs = append(c.Errs, ErrorInfo{E_ERROR, "[翻译错误]:翻译文件错误"})
		return
	}

	c.Lang = &LangSheet{
		SheetRows: sheet,
		FieldRef:  make(map[string]int),
		IdRef:     make(map[string]int),
	}
	for i, row := range sheet {
		if i == 0 { //第一行
			for j, text := range row {
				if strings.Contains(text, "_翻译") {
					fieldName := strings.TrimRight(text, "_翻译")
					c.Lang.FieldRef[fieldName] = j
				}
			}
		} else {
			c.Lang.IdRef[row[0]] = i
		}
	}
}

func (c *XlsxConv) getLangCellText(id string, f FieldInfo) string {
	if f.Mode != "r" && (f.Type == "table" || f.Type == "object" || f.Type == "string") {
		rId, rOk := c.Lang.IdRef[id]
		cId, cOk := c.Lang.FieldRef[f.Name]
		if rOk && cOk {
			return c.Lang.SheetRows[rId][cId]
		}
	}
	return ""
}

func (c *XlsxConv) checkLangText(langText, base string, id, f string) bool {
	if len(langText) > 0 {
		flags := mainForm.TestChkBox.Checked()
		if flags && !checkAscii(base, langText) {
			errStr := fmt.Sprintf("[翻译内容不匹配 id=%s,字段=%s]:源=%s,翻译=%s", id, f, base, langText)
			c.Errs = append(c.Errs, ErrorInfo{E_ERROR, errStr})
			return false
		}

		if c.FileName == "string.xlsx" || c.FileName == "error.xlsx" {
			re := regexp.MustCompile(`%[a-z]`)
			reSlice1 := re.FindAllString(base, -1)
			reSlice2 := re.FindAllString(langText, -1)
			if len(reSlice1) != len(reSlice2) {
				c.Errs = append(c.Errs, ErrorInfo{E_ERROR, fmt.Sprintf("占位符个数不匹配 id=%s,字段=%s]", id, f)})
			}
		}
		return true
	} else {
		if isChineseChar(base) {
			c.Errs = append(c.Errs, ErrorInfo{E_WARN, fmt.Sprintf("[翻译缺失 id=%s,字段=%s]", id, f)})
		}
	}
	return false
}

func (c *XlsxConv) loadXlsxHead(workSheet [][]string) {
	if len(workSheet) < 4 {
		c.Errs = append(c.Errs, ErrorInfo{E_ERROR, "配置头至少要4行"})
		return
	}

	fieldRow := workSheet[1] //字段名
	typeRow := workSheet[2]  //字段类型
	modeRow := workSheet[3]  //生成方式
	if len(fieldRow) == 0 {
		c.Errs = append(c.Errs, ErrorInfo{E_ERROR, "[配置头错误]:配置字段名为空"})
		return
	}

	c.Fields = make(map[int]FieldInfo, 50)
	for i, fieldName := range fieldRow {
		fieldType := typeRow[i]
		modeType := modeRow[i]
		if modeType == "s" || modeType == "d" {
			c.hasSrvField = true
		}

		// 检测key字段
		if i == 0 { //字段行的第一个字段为配置的key,需要检查下
			if len(fieldName) == 0 {
				c.Errs = append(c.Errs, ErrorInfo{E_ERROR, "[配置头错误]:配置没有key"})
			}
			if fieldType != "int" && fieldType != "string" {
				c.Errs = append(c.Errs, ErrorInfo{E_ERROR, "[配置头错误]:id字段类型错误"})
			}
		}

		// 字段通用检查
		if modeType == "r" {
			continue
		} else {
			var errStr string
			if len(fieldName) > 0 {
				if strings.Contains(fieldName, " ") {
					errStr = fmt.Sprintf("[配置头错误]:字段[%s]有空格", fieldName)
					c.Errs = append(c.Errs, ErrorInfo{E_ERROR, errStr})
				}
				if modeType != "c" && modeType != "s" && modeType != "d" {
					errStr = fmt.Sprintf("[配置头错误]:字段[%s]生成方式错误", fieldName)
					c.Errs = append(c.Errs, ErrorInfo{E_ERROR, errStr})
				}
				if len(fieldType) == 0 {
					errStr = fmt.Sprintf("[配置头错误]:字段[%s]类型不存在", fieldName)
					c.Errs = append(c.Errs, ErrorInfo{E_ERROR, errStr})
				}
			} else {
				if len(modeType) > 0 || len(fieldType) > 0 {
					errStr = fmt.Sprintf("[配置头错误]:第[%d]个字段名为空", i+1)
					c.Errs = append(c.Errs, ErrorInfo{E_ERROR, errStr})
				}
			}
			if len(errStr) == 0 {
				c.Fields[i] = FieldInfo{fieldName, fieldType, modeType}
			}
		}
	}
}

func (c *XlsxConv) generate() {
	startTime := time.Now()
	c.Errs = make([]ErrorInfo, 0, 5)
	defer func() {
		if err := recover(); err != nil {
			c.Errs = append(c.Errs, ErrorInfo{E_ERROR, fmt.Sprintf("%v", err)})
		}
		c.Msec = int(time.Now().Sub(startTime).Nanoseconds() / 1e6)
		ConvChan <- c
	}()

	xlFile, err := excelize.OpenFile(c.AbsPath)
	if err != nil {
		c.Errs = append(c.Errs, ErrorInfo{E_ERROR, err.Error()})
		return
	}

	// 第一个sheet为配置
	sheetName := xlFile.GetSheetName(1)
	workSheet := xlFile.GetRows(sheetName) //sheetName = "Sheet1"
	c.loadXlsxHead(workSheet)
	if c.hasError(E_WARN) {
		return
	}

	// 翻译文件
	if c.hasSrvField {
		c.loadLangXlsx()
	}

	// lua
	outFormat := "lua"
	if outFormat == "lua" {
		c.parseToLua(workSheet)
	} else {
		// other,eg:json,yml....
		// TODO
	}
}

func (c *XlsxConv) outPutToFile(rowsSlice []string, format string) {
	outDir := mainForm.getOutPutDir() + "\\" + c.FolderName
	_, err := os.Stat(outDir)
	if os.IsNotExist(err) {
		err = os.MkdirAll(outDir, os.ModePerm)
		if err != nil {
			c.Errs = append(c.Errs, ErrorInfo{E_ERROR, err.Error()})
			return
		}
	}

	name := strings.TrimSuffix(c.FileName, ".xlsx")
	file := fmt.Sprintf("%s%s.%s", outDir, name, format)
	outFile, operr := os.OpenFile(file, os.O_CREATE|os.O_TRUNC|os.O_RDWR, 0666)
	if operr != nil {
		c.Errs = append(c.Errs, ErrorInfo{E_ERROR, operr.Error()})
		return
	}
	defer outFile.Close()

	outFile.WriteString(strings.Join(rowsSlice, "\n"))
	outFile.Sync()
}

// 开始生成
func startConv(idxs map[int]bool) {
	count := len(idxs)
	if count > 0 {
		startTime := time.Now()
		ConvChan = make(chan *XlsxConv, count)
		for idx, _ := range idxs {
			conv := Convs[idx]
			go conv.generate()
		}

		for i := 0; i < count; i++ {
			<-ConvChan
			mainForm.updateProcess()
		}
		saveConvTime(idxs)
		mainForm.ConvResult(idxs, startTime)
	}
}

func saveConvTime(idxs map[int]bool) {
	parentDir := mainForm.getParentDir()
	if len(parentDir) == 0 {
		mainForm.MsgBox("保存xlsx时间出错", "错误")
		return
	}

	file := parentDir + "\\" + "lastModTime.txt"
	outFile, operr := os.OpenFile(file, os.O_CREATE|os.O_TRUNC|os.O_RDWR, 0666)
	if operr != nil {
		mainForm.MsgBox("创建[lastModTime.txt]文件出错", "错误")
	}
	defer outFile.Close()

	modTimes := make([]string, 0, len(Convs))
	for idx, c := range Convs {
		_, ok := idxs[idx]
		if ok && !c.hasError(E_ERROR) {
			modTimes = append(modTimes, c.RelPath+"|"+strconv.FormatUint(c.ModTime, 10))
		} else {
			modTimes = append(modTimes, c.RelPath+"|"+strconv.FormatUint(c.RecordTime, 10))
		}
	}

	outFile.WriteString(strings.Join(modTimes, "\n"))
	outFile.Sync()
}
