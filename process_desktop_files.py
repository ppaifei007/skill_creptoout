import os
import glob
import subprocess
import shutil
from pathlib import Path
from ppt_desensitizer import GPTEnhancedDesensitizer

def extract_rar(rar_path, extract_dir):
    """尝试使用本机 WinRAR 解压 RAR 文件"""
    winrar_paths = [
        r"C:\Program Files\WinRAR\WinRAR.exe",
        r"C:\Program Files (x86)\WinRAR\WinRAR.exe"
    ]
    
    winrar_exe = None
    for p in winrar_paths:
        if os.path.exists(p):
            winrar_exe = p
            break
            
    if winrar_exe:
        print(f"找到 WinRAR: {winrar_exe}，正在解压 {rar_path}...")
        cmd = [winrar_exe, "x", "-y", rar_path, f"{extract_dir}\\"]
        subprocess.run(cmd, check=True)
        return True
    else:
        print(f"⚠️ 未找到 WinRAR，无法自动解压: {rar_path}")
        print(f"👉 请手动解压后，再次运行此脚本，或者把解压后的 pptx 文件直接放到桌面上。")
        return False

def process_desktop():
    desktop_dir = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    output_dir = os.path.join(desktop_dir, '脱敏后输出')
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    print(f"正在加载 GPT-5.4 增强脱敏引擎及语义规则...")
    d = GPTEnhancedDesensitizer('desensitization_config.json')
    
    # 1. 查找桌面上独立的 PPTX 文件
    ppt_files = glob.glob(os.path.join(desktop_dir, '*.pptx'))
    print(f"\n✅ 找到 {len(ppt_files)} 个独立的 PPTX 文件。")
    
    for ppt in ppt_files:
        filename = os.path.basename(ppt)
        out_path = os.path.join(output_dir, f"[已脱敏]_{filename}")
        print(f"⏳ 正在处理: {filename} ...")
        success = d.process_ppt_file(ppt, out_path)
        if success:
            print(f"✅ 处理成功 -> {out_path}")
        else:
            print(f"❌ 处理失败 -> {filename}")
        
    # 2. 查找并解压 RAR 文件
    rar_files = glob.glob(os.path.join(desktop_dir, '*.rar'))
    print(f"\n✅ 找到 {len(rar_files)} 个 RAR 压缩包。")
    
    for rar in rar_files:
        filename = os.path.basename(rar)
        extract_folder = os.path.join(desktop_dir, f"temp_{filename}_extracted")
        if not os.path.exists(extract_folder):
            os.makedirs(extract_folder)
            
        success = extract_rar(rar, extract_folder)
        if success:
            # 遍历解压出来的 PPT
            for root, dirs, files in os.walk(extract_folder):
                for f in files:
                    if f.endswith('.pptx'):
                        in_ppt = os.path.join(root, f)
                        rel_path = os.path.relpath(in_ppt, extract_folder)
                        safe_name = rel_path.replace('\\', '_').replace('/', '_')
                        out_ppt = os.path.join(output_dir, f"[已脱敏]_{filename}_{safe_name}")
                        print(f"⏳ 正在处理压缩包内文件: {f} ...")
                        res = d.process_ppt_file(in_ppt, out_ppt)
                        if res:
                            print(f"✅ 处理成功 -> {out_ppt}")
                        else:
                            print(f"❌ 处理失败 -> {f}")
                        
            # 清理临时文件夹
            shutil.rmtree(extract_folder)

if __name__ == '__main__':
    process_desktop()
    print("\n🎉 所有脱敏任务处理完成！请查看桌面上的 '脱敏后输出' 文件夹。")
