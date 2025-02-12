package main

import (
	"fmt"
	"os"
	"os/exec"
	"strings"
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows/registry"
)

const (
	regPath          = `Directory\shell\AddToPath`
	regPathRemove    = `Directory\shell\RemoveFromPath`
	scriptName       = "add_to_path.vbs"      // 删除
	scriptNameRemove = "remove_from_path.vbs" // 删除
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

	// 更新系统通知
	if err := notifyEnvironmentChange(); err != nil {
		return fmt.Errorf("通知系统环境变量更改失败: %v", err)
	}

	showMessageBox("操作成功", fmt.Sprintf("成功将 %s 添加到环境变量", folderPath))
	os.Exit(0)
	return nil
}

func removeFromPath(folderPath string) error {
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

	currentPath, _, err := k.GetStringValue("Path")
	if err != nil {
		return fmt.Errorf("读取PATH失败: %v", err)
	}

	paths := make(map[string]bool)
	for _, path := range strings.Split(currentPath, ";") {
		path = strings.TrimSpace(path)
		if path != "" && strings.ToLower(path) != strings.ToLower(folderPath) {
			paths[path] = true
		}
	}

	var newPaths []string
	for path := range paths {
		newPaths = append(newPaths, path)
	}
	newPath := strings.Join(newPaths, ";")

	err = k.SetStringValue("Path", newPath)
	if err != nil {
		return fmt.Errorf("更新系统PATH失败: %v", err)
	}

	// 更新系统通知
	if err := notifyEnvironmentChange(); err != nil {
		return fmt.Errorf("通知系统环境变量更改失败: %v", err)
	}

	showMessageBox("操作成功", fmt.Sprintf("成功从环境变量移除 %s", folderPath))
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
	exec.Command("cmd", "/C", "echo %PATH%").Output()

	if duplicateCount > 0 {
		fmt.Printf("成功清理了 %d 个重复的路径\n", duplicateCount)
	} else {
		fmt.Println("未发现重复的路径")
	}
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
	// 使用 rundll32 来发送系统消息
	cmd := exec.Command("rundll32.exe", "user32.dll,UpdatePerUserSystemParameters")
	if err := cmd.Run(); err != nil {
		return fmt.Errorf("UpdatePerUserSystemParameters 失败: %v", err)
	}

	// 广播环境变量更改消息
	cmd = exec.Command("cmd", "/C", "echo %PATH%")
	return cmd.Run()
}
