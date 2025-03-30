package main

import (
	"fmt"
	"os"
	"os/exec"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"golang.org/x/sys/windows/registry"
)

const (
	regPath        = `Directory\shell\AddToPath`
	regPathRemove  = `Directory\shell\RemoveFromPath`
	msgNeedRestart = "请注意：环境变量已更新，但必须重新打开命令窗口才能生效。\n要立即应用更改，请点击\"是\"来关闭所有cmd窗口。"
)

func main() {
	if len(os.Args) < 2 {
		for {
			fmt.Println("\n请选择操作：")
			fmt.Println("1. 安装程序")
			fmt.Println("2. 卸载程序")
			fmt.Println("3. 清理重复的环境变量")
			fmt.Println("4. 退出")

			var choice string
			fmt.Print("请输入选项 (1-4): ")
			fmt.Scanln(&choice)

			switch choice {
			case "1":
				install()
			case "2":
				uninstall()
			case "3":
				cleanupDuplicatePaths()
			case "4":
				return
			default:
				fmt.Println("无效的选项，请重新选择")
			}
		}
	}

	// 处理命令行参数
	switch os.Args[1] {
	case "install":
		install()
	case "uninstall":
		uninstall()
	case "cleanup":
		cleanupDuplicatePaths()
	case "add":
		// 处理右键菜单"添加到环境变量"的情况
		if len(os.Args) >= 3 {
			folderPath := os.Args[2]
			err := addToPath(folderPath)
			if err != nil {
				fmt.Printf("添加到环境变量失败: %v\n", err)
			} else {
				fmt.Printf("成功将 %s 添加到环境变量\n", folderPath)
			}
		} else {
			fmt.Println("缺少文件夹路径参数")
		}
	case "remove":
		// 处理右键菜单"从环境变量移除"的情况
		if len(os.Args) >= 3 {
			folderPath := os.Args[2]
			err := removeFromPath(folderPath)
			if err != nil {
				fmt.Printf("从环境变量移除失败: %v\n", err)
			} else {
				fmt.Printf("成功从环境变量移除 %s\n", folderPath)
			}
		} else {
			fmt.Println("缺少文件夹路径参数")
		}
	default:
		printUsage()
	}

	waitForInput()
}

func waitForInput() {
	fmt.Println("\n按回车键退出...")
	fmt.Scanln() // 等待用户输入
}

func showMessageBox(title, message string) {
	user32 := syscall.NewLazyDLL("user32.dll")
	messageBox := user32.NewProc("MessageBoxW")
	messageBox.Call(0,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(message))),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(title))),
		0x40) // MB_ICONINFORMATION
}

func install() {
	// 添加"添加到环境变量"注册表项
	k, _, err := registry.CreateKey(registry.CLASSES_ROOT, regPath, registry.ALL_ACCESS)
	if err != nil {
		fmt.Printf("创建添加注册表键失败: %v\n", err)
		return
	}
	defer k.Close()

	// 设置默认值
	err = k.SetStringValue("", "添加到环境变量")
	if err != nil {
		fmt.Printf("设置注册表值失败: %v\n", err)
		return
	}

	// 设置图标
	err = k.SetStringValue("Icon", `C:\Windows\System32\shell32.dll,177`)
	if err != nil {
		fmt.Printf("设置图标失败: %v\n", err)
		return
	}

	// 创建command子键
	cmdKey, _, err := registry.CreateKey(registry.CLASSES_ROOT, regPath+`\command`, registry.ALL_ACCESS)
	if err != nil {
		fmt.Printf("创建command键失败: %v\n", err)
		return
	}
	defer cmdKey.Close()

	// 获取程序的完整路径
	exePath, err := os.Executable()
	if err != nil {
		fmt.Printf("获取程序路径失败: %v\n", err)
		return
	}

	// 使用完整路径设置命令
	command := fmt.Sprintf(`"%s" add "%%1"`, exePath)
	err = cmdKey.SetStringValue("", command)
	if err != nil {
		fmt.Printf("设置命令失败: %v\n", err)
		return
	}

	// 添加"从环境变量移除"注册表项
	kRemove, _, err := registry.CreateKey(registry.CLASSES_ROOT, regPathRemove, registry.ALL_ACCESS)
	if err != nil {
		fmt.Printf("创建移除注册表键失败: %v\n", err)
		return
	}
	defer kRemove.Close()

	// 设置默认值
	err = kRemove.SetStringValue("", "从环境变量移除")
	if err != nil {
		fmt.Printf("设置注册表值失败: %v\n", err)
		return
	}

	// 设置图标
	err = kRemove.SetStringValue("Icon", `C:\Windows\System32\shell32.dll,132`)
	if err != nil {
		fmt.Printf("设置图标失败: %v\n", err)
		return
	}

	// 创建command子键
	cmdKeyRemove, _, err := registry.CreateKey(registry.CLASSES_ROOT, regPathRemove+`\command`, registry.ALL_ACCESS)
	if err != nil {
		fmt.Printf("创建command键失败: %v\n", err)
		return
	}
	defer cmdKeyRemove.Close()

	// 使用完整路径设置移除命令
	commandRemove := fmt.Sprintf(`"%s" remove "%%1"`, exePath)
	err = cmdKeyRemove.SetStringValue("", commandRemove)
	if err != nil {
		fmt.Printf("设置命令失败: %v\n", err)
		return
	}

	showMessageBox("安装成功", "右键菜单已成功安装！")
	os.Exit(0)
}

func uninstall() {
	// 删除注册表项（添加到环境变量）
	err := registry.DeleteKey(registry.CLASSES_ROOT, regPath+`\command`)
	if err != nil && err != registry.ErrNotExist {
		fmt.Printf("删除command键失败: %v\n", err)
		return
	}

	err = registry.DeleteKey(registry.CLASSES_ROOT, regPath)
	if err != nil && err != registry.ErrNotExist {
		fmt.Printf("删除注册表键失败: %v\n", err)
		return
	}

	// 删除注册表项（从环境变量移除）
	err = registry.DeleteKey(registry.CLASSES_ROOT, regPathRemove+`\command`)
	if err != nil && err != registry.ErrNotExist {
		fmt.Printf("删除移除command键失败: %v\n", err)
		return
	}

	err = registry.DeleteKey(registry.CLASSES_ROOT, regPathRemove)
	if err != nil && err != registry.ErrNotExist {
		fmt.Printf("删除移除注册表键失败: %v\n", err)
		return
	}

	showMessageBox("卸载成功", "右键菜单已成功卸载！")
	os.Exit(0)
}

func addToPath(folderPath string) error {
	// 检查是否具有管理员权限
	if !isAdmin() {
		return fmt.Errorf("需要管理员权限才能修改系统环境变量")
	}

	k, err := registry.OpenKey(registry.LOCAL_MACHINE,
		`SYSTEM\CurrentControlSet\Control\Session Manager\Environment`,
		registry.QUERY_VALUE|registry.SET_VALUE)
	if err != nil {
		return fmt.Errorf("打开系统环境变量失败: %v", err)
	}
	defer k.Close()

	// 读取当前PATH
	currentPath, _, err := k.GetStringValue("Path")
	if err != nil {
		return fmt.Errorf("读取PATH失败: %v", err)
	}

	// 分割并去重
	paths := make(map[string]bool)
	for _, path := range strings.Split(currentPath, ";") {
		path = strings.TrimSpace(path)
		if path != "" {
			paths[strings.ToLower(path)] = true
		}
	}

	// 添加新路径
	paths[strings.ToLower(folderPath)] = true

	// 构建新的PATH
	var newPaths []string
	for path := range paths {
		newPaths = append(newPaths, path)
	}
	newPath := strings.Join(newPaths, ";")

	// 更新系统PATH
	err = k.SetStringValue("Path", newPath)
	if err != nil {
		return fmt.Errorf("更新系统PATH失败: %v", err)
	}

	// 通知系统环境变量已更改
	notifyEnvironmentChange()

	// 询问用户是否要关闭所有CMD窗口以使更改立即生效
	if askYesNo(msgNeedRestart) {
		killCmdProcesses()
	}

	showMessageBox("操作成功", fmt.Sprintf("成功将 %s 添加到环境变量\n\n若未关闭命令窗口，请记得重新打开命令窗口才能使变更生效", folderPath))
	os.Exit(0)
	return nil
}

func removeFromPath(folderPath string) error {
	// 检查是否具有管理员权限
	if !isAdmin() {
		return fmt.Errorf("需要管理员权限才能修改系统环境变量")
	}

	// 添加调试输出
	fmt.Printf("准备移除的路径: %s\n", folderPath)
	fmt.Printf("准备移除的路径(小写): %s\n", strings.ToLower(folderPath))

	k, err := registry.OpenKey(registry.LOCAL_MACHINE,
		`SYSTEM\CurrentControlSet\Control\Session Manager\Environment`,
		registry.QUERY_VALUE|registry.SET_VALUE)
	if err != nil {
		return fmt.Errorf("打开系统环境变量失败: %v", err)
	}
	defer k.Close()

	currentPath, _, err := k.GetStringValue("Path")
	if err != nil {
		return fmt.Errorf("读取PATH失败: %v", err)
	}

	// 添加调试输出
	fmt.Println("\n当前PATH包含以下路径:")
	for i, path := range strings.Split(currentPath, ";") {
		if path != "" {
			fmt.Printf("%d: [%s] (小写: [%s])\n", i, path, strings.ToLower(path))
		}
	}

	// 将folderPath转为小写以便比较
	lowerFolderPath := strings.ToLower(folderPath)

	// 创建一个记录需要保留的路径的map
	keepPaths := make([]string, 0)
	removedPaths := make([]string, 0)

	for _, path := range strings.Split(currentPath, ";") {
		pathTrimmed := strings.TrimSpace(path)
		if pathTrimmed == "" {
			continue
		}

		// 比较小写版本
		lowerPath := strings.ToLower(pathTrimmed)

		if lowerPath != lowerFolderPath {
			keepPaths = append(keepPaths, pathTrimmed)
		} else {
			removedPaths = append(removedPaths, pathTrimmed)
			fmt.Printf("找到匹配项将被移除: [%s]\n", pathTrimmed)
		}
	}

	// 添加调试信息
	if len(removedPaths) == 0 {
		fmt.Println("警告: 未找到要移除的路径!")
		fmt.Printf("要移除的路径(小写): [%s]\n", lowerFolderPath)
	}

	newPath := strings.Join(keepPaths, ";")

	// 添加调试输出
	fmt.Println("\n更新后的PATH将包含以下路径:")
	for i, path := range keepPaths {
		fmt.Printf("%d: [%s]\n", i, path)
	}

	fmt.Println("\n以下路径将被移除:")
	for i, path := range removedPaths {
		fmt.Printf("%d: [%s]\n", i, path)
	}

	// 更新系统PATH
	err = k.SetStringValue("Path", newPath)
	if err != nil {
		return fmt.Errorf("更新系统PATH失败: %v", err)
	}

	// 通知系统环境变量已更改
	notifyEnvironmentChange()

	// 构建消息
	successMsg := fmt.Sprintf("成功从环境变量移除 %s", folderPath)
	if len(removedPaths) == 0 {
		successMsg = fmt.Sprintf("警告：未找到要移除的路径 %s\n可能是路径格式或大小写不匹配", folderPath)
		showMessageBox("操作完成", successMsg)
	} else {
		// 询问用户是否要关闭所有CMD窗口以使更改立即生效
		if askYesNo(msgNeedRestart) {
			killCmdProcesses()
			successMsg += "\n\n已尝试关闭所有命令窗口，打开新窗口即可应用更改"
		} else {
			successMsg += "\n\n请记得重新打开命令窗口才能使变更生效"
		}
		// showMessageBox("操作成功", successMsg)
	}

	os.Exit(0)
	return nil
}

func cleanupDuplicatePaths() {
	k, err := registry.OpenKey(registry.LOCAL_MACHINE,
		`SYSTEM\CurrentControlSet\Control\Session Manager\Environment`,
		registry.QUERY_VALUE|registry.SET_VALUE)
	if err != nil {
		fmt.Printf("打开系统环境变量失败: %v\n", err)
		return
	}
	defer k.Close()

	currentPath, _, err := k.GetStringValue("Path")
	if err != nil {
		fmt.Printf("读取PATH失败: %v\n", err)
		return
	}

	paths := make(map[string]bool)
	duplicateCount := 0
	for _, path := range strings.Split(currentPath, ";") {
		path = strings.TrimSpace(path)
		if path != "" {
			if paths[strings.ToLower(path)] {
				duplicateCount++
			}
			paths[strings.ToLower(path)] = true
		}
	}

	var newPaths []string
	for path := range paths {
		newPaths = append(newPaths, path)
	}
	newPath := strings.Join(newPaths, ";")

	err = k.SetStringValue("Path", newPath)
	if err != nil {
		fmt.Printf("更新系统PATH失败: %v\n", err)
		return
	}

	// 通知系统环境变量已更改
	notifyEnvironmentChange()

	// 构建消息
	var message string
	if duplicateCount > 0 {
		message = fmt.Sprintf("成功清理了 %d 个重复的路径", duplicateCount)

		// 询问用户是否要关闭所有CMD窗口以使更改立即生效
		if askYesNo(msgNeedRestart) {
			killCmdProcesses()
			message += "\n\n已尝试关闭所有命令窗口，打开新窗口即可应用更改"
		} else {
			message += "\n\n请记得重新打开命令窗口才能使变更生效"
		}
	} else {
		message = "未发现重复的路径"
	}

	showMessageBox("清理结果", message)
}

func printUsage() {
	fmt.Println("使用方法:")
	fmt.Println("  install   - 安装程序")
	fmt.Println("  uninstall - 卸载程序")
	fmt.Println("  cleanup   - 清理重复的环境变量")
	fmt.Println("  add <path>   - 添加指定路径到环境变量")
	fmt.Println("  remove <path> - 从环境变量移除指定路径")
}

// 新增辅助函数
func isAdmin() bool {
	_, err := os.Open("\\\\.\\PHYSICALDRIVE0")
	return err == nil
}

func notifyEnvironmentChange() error {
	// 打印提示
	fmt.Println("环境变量已更新，尝试刷新系统缓存...")

	// 1. 使用rundll32来刷新环境变量（异步执行）
	go exec.Command("rundll32.exe", "user32.dll,UpdatePerUserSystemParameters").Run()

	// 2. 直接使用Windows API广播环境变量更改消息
	user32 := syscall.NewLazyDLL("user32.dll")
	sendMessageTimeout := user32.NewProc("SendMessageTimeoutW")

	// 使用SendMessageTimeout而不是PostMessage，确保消息被处理
	sendMessageTimeout.Call(
		0xFFFF, // HWND_BROADCAST
		0x001A, // WM_SETTINGCHANGE
		0,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr("Environment"))),
		0x0002, // SMTO_ABORTIFHUNG
		1000,   // 超时时间（毫秒）
		0,      // 结果（不关心）
	)

	// 3. 模拟环境变量编辑器的保存操作
	// 使用PowerShell来刷新环境变量（这类似于在GUI中保存环境变量）
	psScript := `
	$objShell = New-Object -ComObject WScript.Shell
	$objEnvironment = $objShell.Environment("SYSTEM")
	# 读取当前PATH并重新设置（模拟保存操作）
	$path = [Environment]::GetEnvironmentVariable("Path", "Machine")
	[Environment]::SetEnvironmentVariable("Path", $path, "Machine")
	Write-Host "环境变量已刷新"
	`

	// 将PowerShell脚本保存到临时文件
	tmpFile, err := os.CreateTemp("", "refresh-env-*.ps1")
	if err == nil {
		defer os.Remove(tmpFile.Name())
		tmpFile.WriteString(psScript)
		tmpFile.Close()

		// 以管理员权限执行PowerShell脚本
		go exec.Command("powershell", "-ExecutionPolicy", "Bypass", "-File", tmpFile.Name()).Run()
	}

	// 4. 使用SETX命令再次刷新PATH变量
	// 这会强制Windows更新环境变量缓存
	k, err := registry.OpenKey(registry.LOCAL_MACHINE,
		`SYSTEM\CurrentControlSet\Control\Session Manager\Environment`,
		registry.QUERY_VALUE)
	if err == nil {
		defer k.Close()
		if currentPath, _, err := k.GetStringValue("Path"); err == nil {
			// 创建一个临时的批处理文件来执行SETX
			batchFile, err := os.CreateTemp("", "refresh-path-*.bat")
			if err == nil {
				defer os.Remove(batchFile.Name())
				batchFile.WriteString(fmt.Sprintf("@echo off\nsetx PATH \"%s\" /M\n", currentPath))
				batchFile.Close()

				// 异步执行批处理文件
				go exec.Command(batchFile.Name()).Run()
			}
		}
	}

	fmt.Println("环境变量已更新，请重新打开命令窗口才能生效")
	return nil
}

// 添加询问用户是否/否的辅助函数
func askYesNo(message string) bool {
	user32 := syscall.NewLazyDLL("user32.dll")
	messageBox := user32.NewProc("MessageBoxW")
	result, _, _ := messageBox.Call(0,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(message))),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr("确认"))),
		0x00000004) // MB_YESNO

	return result == 6 // IDYES = 6
}

// 添加尝试关闭所有CMD窗口的函数
func killCmdProcesses() {
	fmt.Println("正在关闭所有命令窗口...")
	// 异步执行，不等待结果
	go exec.Command("taskkill", "/F", "/IM", "cmd.exe").Run()
	go exec.Command("taskkill", "/F", "/IM", "powershell.exe").Run()
	// 等待一小段时间，让进程有时间结束
	time.Sleep(500 * time.Millisecond)
}
