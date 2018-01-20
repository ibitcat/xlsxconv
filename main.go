package main

import (
	"gitee.com/ying32/govcl/vcl"
)

var (
	mainForm *TFormConv
	Convs    []*XlsxConv
	ConvChan chan *XlsxConv
)

func main() {
	mainForm = CreateMainForm()
	mainForm.CreateControl()
	mainForm.LoadXlxs()

	// app run
	vcl.Application.Run()
	mainForm.Inifile.Free()
}
