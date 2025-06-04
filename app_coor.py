#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# author: Eddie
"""
使用pywinauto连接app并截图标出所有的组件
"""

import os
import time
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import pywinauto
from pywinauto import Application, Desktop
from pywinauto.controls.uiawrapper import UIAWrapper
import traceback


class AppCoordinator:
    def __init__(self):
        self.app = None
        self.main_window = None
        self.screenshot_dir = "screenshots"
        self.ensure_screenshot_dir()
    
    def ensure_screenshot_dir(self):
        """确保截图目录存在"""
        if not os.path.exists(self.screenshot_dir):
            os.makedirs(self.screenshot_dir)
    
    def list_running_applications(self):
        """列出当前运行的所有应用程序"""
        print("正在获取运行中的应用程序...")
        try:
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            
            print(f"\n找到 {len(windows)} 个窗口:")
            for i, window in enumerate(windows):
                try:
                    title = window.window_text()
                    class_name = window.class_name()
                    process_id = window.process_id()
                    if title.strip():  # 只显示有标题的窗口
                        print(f"{i+1}. 标题: '{title}' | 类名: '{class_name}' | 进程ID: {process_id}")
                except:
                    continue
            return windows
        except Exception as e:
            print(f"获取应用程序列表失败: {str(e)}")
            return []
    
    def connect_to_app(self, app_name=None, title=None, process_id=None):
        """连接到指定的应用程序"""
        try:
            if process_id:
                self.app = Application(backend="uia").connect(process=process_id)
            elif title:
                self.app = Application(backend="uia").connect(title=title)
            elif app_name:
                self.app = Application(backend="uia").connect(path=app_name)
            else:
                # 如果没有指定，尝试连接到计算器作为示例
                self.app = Application(backend="uia").start("calc.exe")
                time.sleep(2)
            
            # 获取主窗口
            if self.app.windows():
                self.main_window = self.app.windows()[0]
                print(f"成功连接到应用程序: {self.main_window.window_text()}")
                return True
            else:
                print("找不到应用程序窗口")
                return False
                
        except Exception as e:
            print(f"连接应用程序失败: {str(e)}")
            return False
    
    def get_all_controls(self, window=None):
        """递归获取所有控件"""
        if window is None:
            window = self.main_window
        
        controls = []
        
        def traverse_controls(control, depth=0):
            try:
                # 获取控件信息
                rect = control.rectangle()
                control_info = {
                    'control': control,
                    'rect': rect,
                    'text': control.window_text(),
                    'class_name': control.class_name(),
                    'control_type': getattr(control, 'control_type', 'Unknown'),
                    'depth': depth
                }
                controls.append(control_info)
                
                # 递归遍历子控件
                try:
                    children = control.children()
                    for child in children:
                        traverse_controls(child, depth + 1)
                except:
                    pass
                    
            except Exception as e:
                print(f"遍历控件时出错: {e}")
        
        try:
            traverse_controls(window)
        except Exception as e:
            print(f"获取控件失败: {e}")
        
        return controls
    
    def take_screenshot_with_annotations(self, save_path=None):
        """截图并标注所有组件"""
        if not self.main_window:
            print("没有连接到应用程序")
            return None
        
        try:
            # 确保窗口在前台
            self.main_window.set_focus()
            time.sleep(0.5)
            
            # 获取窗口截图
            screenshot = self.main_window.capture_as_image()
            
            # 获取所有控件
            controls = self.get_all_controls()
            print(f"找到 {len(controls)} 个控件")
            
            # 在截图上标注控件
            draw = ImageDraw.Draw(screenshot)
            
            # 尝试加载字体
            try:
                font = ImageDraw.getfont()
            except:
                font = None
            
            # 为不同类型的控件使用不同颜色
            colors = {
                'Button': 'red',
                'Edit': 'blue',
                'Static': 'green',
                'ListBox': 'purple',
                'ComboBox': 'orange',
                'TreeView': 'brown',
                'TabControl': 'pink',
                'Default': 'red'
            }
            
            window_rect = self.main_window.rectangle()
            
            for i, control_info in enumerate(controls):
                try:
                    rect = control_info['rect']
                    control_type = control_info['control_type']
                    text = control_info['text']
                    
                    # 计算相对于窗口的坐标
                    x1 = rect.left - window_rect.left
                    y1 = rect.top - window_rect.top
                    x2 = rect.right - window_rect.left
                    y2 = rect.bottom - window_rect.top
                    
                    # 确保坐标在图像范围内
                    if x1 < 0 or y1 < 0 or x2 > screenshot.width or y2 > screenshot.height:
                        continue
                    
                    if x2 - x1 <= 1 or y2 - y1 <= 1:  # 跳过太小的控件
                        continue
                    
                    # 选择颜色
                    color = colors.get(control_type, colors['Default'])
                    
                    # 绘制边框
                    draw.rectangle([x1, y1, x2, y2], outline=color, width=2)
                    
                    # 绘制控件编号
                    draw.text((x1 + 2, y1 + 2), str(i+1), fill=color, font=font)
                    
                    print(f"{i+1:3d}. {control_type:15s} | {text[:30]:30s} | ({x1},{y1},{x2},{y2})")
                    
                except Exception as e:
                    print(f"标注控件 {i} 时出错: {e}")
                    continue
            
            # 保存截图
            if save_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                app_name = self.main_window.window_text().replace(' ', '_').replace('/', '_')
                save_path = os.path.join(self.screenshot_dir, f"{app_name}_{timestamp}.png")
            
            screenshot.save(save_path)
            print(f"\n截图已保存到: {save_path}")
            
            return save_path, controls
            
        except Exception as e:
            print(f"截图失败: {str(e)}")
            traceback.print_exc()
            return None, []
    
    def print_control_tree(self, controls):
        """打印控件树结构"""
        print("\n控件树结构:")
        print("-" * 80)
        for i, control_info in enumerate(controls):
            depth = control_info['depth']
            indent = "  " * depth
            text = control_info['text'][:20] if control_info['text'] else "(无文本)"
            control_type = control_info['control_type']
            class_name = control_info['class_name']
            
            print(f"{i+1:3d}. {indent}{control_type} | {class_name} | '{text}'")
    
    def interactive_mode(self):
        """交互模式"""
        print("=== PyWinAuto App 截图标注工具 ===\n")
        
        while True:
            print("\n请选择操作:")
            print("1. 列出运行中的应用程序")
            print("2. 连接到应用程序")
            print("3. 截图并标注控件")
            print("4. 打印控件树")
            print("5. 启动计算器(示例)")
            print("0. 退出")
            
            choice = input("\n请输入选择 (0-5): ").strip()
            
            if choice == '0':
                break
            elif choice == '1':
                self.list_running_applications()
            elif choice == '2':
                self.connect_app_interactive()
            elif choice == '3':
                if self.main_window:
                    save_path, controls = self.take_screenshot_with_annotations()
                    if save_path:
                        self.print_control_tree(controls)
                else:
                    print("请先连接到应用程序")
            elif choice == '4':
                if self.main_window:
                    controls = self.get_all_controls()
                    self.print_control_tree(controls)
                else:
                    print("请先连接到应用程序")
            elif choice == '5':
                print("正在启动计算器...")
                if self.connect_to_app():
                    print("计算器启动成功!")
                    time.sleep(1)
                    save_path, controls = self.take_screenshot_with_annotations()
                    if save_path:
                        self.print_control_tree(controls)
            else:
                print("无效选择，请重试")
    
    def connect_app_interactive(self):
        """交互式连接应用程序"""
        print("\n连接方式:")
        print("1. 通过窗口标题")
        print("2. 通过进程ID")
        print("3. 启动新程序")
        
        method = input("请选择连接方式 (1-3): ").strip()
        
        if method == '1':
            title = input("请输入窗口标题 (支持部分匹配): ").strip()
            if title:
                self.connect_to_app(title=title)
        elif method == '2':
            try:
                pid = int(input("请输入进程ID: ").strip())
                self.connect_to_app(process_id=pid)
            except ValueError:
                print("进程ID必须是数字")
        elif method == '3':
            app_path = input("请输入程序路径 (如: calc.exe): ").strip()
            if app_path:
                self.connect_to_app(app_name=app_path)
        else:
            print("无效选择")


def main():
    """主函数"""
    coordinator = AppCoordinator()
    
    # 检查是否有命令行参数
    import sys
    if len(sys.argv) > 1:
        # 命令行模式
        app_name = sys.argv[1]
        print(f"正在连接到: {app_name}")
        if coordinator.connect_to_app(title=app_name):
            save_path, controls = coordinator.take_screenshot_with_annotations()
            if save_path:
                coordinator.print_control_tree(controls)
    else:
        # 交互模式
        coordinator.interactive_mode()


if __name__ == "__main__":
    main()
