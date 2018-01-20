// 加载xlsx配置文件

package main

import (
	"bufio"
	"io/ioutil"
	"os"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
)

func FindLangFolder(dir string) string {
	if len(dir) > 0 {
		var langDir string
		langDir = dir + "\\language"
		dir, err := ioutil.ReadDir(langDir)
		if err != nil {
			return ""
		}

		var realDir string
		for _, fi := range dir {
			if fi.IsDir() {
				dirRoot := langDir + "\\" + fi.Name()
				filepath.Walk(dirRoot, func(path string, f os.FileInfo, err error) error {
					if !f.IsDir() {
						ok, mErr := filepath.Match("$*.xlsx", f.Name())
						if ok {
							realDir = fi.Name()
							return nil
						}
						return mErr
					}
					return nil
				})
			}
		}

		if len(realDir) > 0 {
			langDir += ("\\" + realDir)
			return langDir
		}
	}
	return ""
}

func WalkXlsx(dir string) error {
	Convs = make([]*XlsxConv, 0, 500)
	_, err := os.Stat(dir)
	notExist := os.IsNotExist(err)
	if notExist {
		return err
	}

	modTime := loadLastModTime(dir)
	_getLastConvTime := func(fileName string) uint64 {
		if modTime != nil {
			tm, ok := modTime[fileName]
			if ok {
				return tm
			}
		}
		return 0
	}

	err = filepath.Walk(dir, func(path string, f os.FileInfo, err error) error {
		ok, mErr := filepath.Match("[^~$]*.xlsx", f.Name())
		if ok {
			if f == nil {
				return err
			}
			if f.IsDir() {
				return nil
			}

			conv := &XlsxConv{
				AbsPath:  path,
				RelPath:  strings.TrimPrefix(path, dir+"\\"),
				FileName: f.Name(),
				ModTime:  uint64(f.ModTime().UnixNano() / 1000000),
			}
			conv.FolderName = strings.TrimSuffix(conv.RelPath, conv.FileName)
			conv.RecordTime = _getLastConvTime(conv.RelPath)
			Convs = append(Convs, conv)
			return nil
		}
		return mErr
	})

	if err != nil {
		return err
	}

	return nil
}

func loadLastModTime(dir string) map[string]uint64 {
	parentDir, err := filepath.Abs(dir + "\\..")
	outDir := parentDir + "\\l-xlsx"
	_, err = os.Stat(outDir)
	notExist := os.IsNotExist(err)
	if notExist { // 输出文件夹不存在
		return nil
	}

	file, ferr := os.Open(parentDir + "\\lastModTime.txt")
	if ferr != nil {
		if runtime.GOOS == "windows" {
			os.RemoveAll(outDir)
		}
		return nil
	}
	defer file.Close()

	lastModTime := make(map[string]uint64, 500)
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := scanner.Text()
		if len(line) > 0 {
			s := strings.Split(line, "|")
			if len(s) == 2 {
				tm, _ := strconv.ParseUint(s[1], 10, 64)
				lastModTime[s[0]] = tm
			}
		}
	}
	return lastModTime
}
