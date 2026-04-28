# Source2Docx-cross-platform

本仓库是原 [Source2Docx](https://github.com/CannonTang/Source2Docx) WPF 工具的 `Avalonia UI` 迁移版本。它保留了原工具的核心目标：自动清洗源代码中的注释与空白行、按指定顺序合并内容，并导出为适合软著材料整理的 `.docx` 文档。

当前仓库是迁移版实现，整体功能、操作方式和原工程保持一致；当前以 `Avalonia` 发布的macos桌面版本为主，并具备继续向多平台发布演进的潜力。

---

## 程序截图

![](image-1.png)

## 功能特点

- 支持按文件顺序导出代码内容到 Word 文档
- 支持批量添加文件、导入文件夹，或以“主文件 + 目录”的方式递归导入
- 同一文件只会被导入一次，重复文件会自动忽略
- 支持勾选 / 取消勾选文件，按需删除
- 支持拖拽调整文件顺序
- 支持多种代码类型
  - `C`
  - `C++`
  - `C#`
  - `C#/WPF`
  - `Java/Eclipse`
  - `Python`
  - `JavaScript/TypeScript`
  - `Go`
  - `Rust`
  - `Kotlin`
  - `Swift`
  - `PHP`
- 支持注释清洗与空白行整理
- 支持前后页裁剪导出

## 当前版本说明

- 原工程：`Windows + WPF`
- 当前工程：`Avalonia UI` 迁移版
- 目标框架：`.NET 10`

除桌面框架与部分平台差异外，迁移版尽量保持原项目的交互与功能逻辑不变。

## 当前导入方式

当前版本支持三种导入方式：

1. `增加文件`
逐个或批量选择文件导入。

2. `导入文件夹`
选择一个目录，递归扫描其内部所有符合当前代码类型的文件并导入。

3. `主文件 + 目录`
先指定一个主文件，再自动导入其所在目录下其他符合条件的文件，并将主文件移动到列表顶部。

## 关于导出裁剪

原 WPF 版本在启用“前 N 页 + 后 N 页裁剪”时，会在文档生成完成后调用 Office 自动化接口，根据 **真实分页结果** 做裁剪，因此页数结果是精确的。

当前迁移版保留了裁剪功能，但实现方式不同：

- 裁剪发生在写入 DOCX 之前
- 程序按“每页约 50 行”的规则预先保留前 `N * 50` 行与后 `N * 50` 行
- 然后再生成最终文档

这意味着：

- 当源码中的每一行都能在文档里保持单行显示时，结果会比较接近预期
- 如果存在超长行、属性声明、长字符串、注解等内容，Word / WPS 在排版时会自动折行
- 一旦折行，最终视觉页数就可能大于 `2N`，但本方法确保了导出的代码行数与源代码一致，此误差是可以接受的

## 使用说明

1. 选择代码类型。
2. 添加单个文件、导入整个文件夹，或通过“主文件 + 目录”递归导入。
3. 根据需要调整右侧源文件列表顺序。
4. 填写软件名称、版本号和输出路径。
5. 如需裁剪导出页数，勾选对应选项并设置 `N` 值。
6. 点击“开始生成”导出文档。

生成文档时，程序会按照右侧源文件列表顺序依次写入代码内容。通常建议将程序入口相关文件放在前面，将收尾或结束逻辑放在最后面。

## 已知限制

- 新增的 `JavaScript/TypeScript`、`Go`、`Rust`、`Kotlin`、`Swift`、`PHP` 目前复用现有 `C` 风格注释清洗器；对常见源码足够实用，但在极端语法场景下仍建议人工复核导出结果。

## 构建

开发环境：

- `JetBrains Rider` 或其他支持 `.NET 10` / `Avalonia` 的 IDE
- 安装 `.NET 10 SDK`
- 任意支持 `.NET 10` 与 `Avalonia` 的桌面开发环境

运行：

```bash
dotnet build
dotnet run --project Source2Docx.csproj
```

## 发布 macOS app

生成 Apple Silicon Mac 的 `.app` 与 `.zip`：

```bash
chmod +x scripts/build-macos-app.sh
./scripts/build-macos-app.sh
```

如需 Intel Mac 版本：

```bash
./scripts/build-macos-app.sh osx-x64
```

生成结果位于 `dist/` 目录：

- `dist/Source2Docx-osx-arm64.app`
- `dist/Source2Docx-osx-arm64.zip`

说明：

- 当前生成的是可直接运行的 `.app`，并带有本地 `ad-hoc` 签名，适合开发测试与小范围分发
- 如果应用是从浏览器下载并解压得到，macOS 可能附加隔离标记，首次打开时会被 Gatekeeper 拦截
- 如遇到“已损坏”或无法打开，可先执行：

```bash
xattr -dr com.apple.quarantine dist/Source2Docx-osx-arm64.app
```

- 若要面向更多用户正式分发，仍建议后续接入 Apple Developer ID 签名与公证
