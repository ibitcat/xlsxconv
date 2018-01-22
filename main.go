package main

import (
	"gitee.com/ying32/govcl/vcl"
)

var (
	ConvForm *TFormConv
	Convs    []*XlsxConv
	ConvChan chan *XlsxConv
)

func main() {
	ConvForm = CreateMainForm()
	ConvForm.CreateControl()
	ConvForm.LoadXlxs()

	// app run
	vcl.Application.Run()
	ConvForm.Inifile.Free()
	ConvForm.icon.Free()
}
