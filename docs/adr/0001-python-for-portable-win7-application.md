---
status: accepted
---

# 使用 Python 构建 Windows 7 便携桌面应用

项目使用与 64 位 Windows 7 SP1 兼容的 64 位 Python 3.8 系列开发，并把解释器和运行依赖一同封装到便携目录中，使客户机复制目录后可直接双击打开 UI，无需安装 Python。选择 Python 是因为本项目以 Excel 读取和文件系统批处理为核心，Python 的实现与维护成本低于 C++ 或 Delphi；代价是必须锁定旧版兼容工具链，并在真实 64 位 Windows 7 SP1 环境完成打包产物验收。

## Considered Options

- C++ 或 Delphi：原生兼容性更强，但开发、Excel 解析和长期维护成本更高。
- C# WinForms：Windows UI 开发便利，但无法假定每台客户机都已具备所需 .NET Framework 版本。
- Python：在开发效率、Excel 生态和可维护性之间最均衡，但 Windows 7 兼容性必须通过固定版本与真实环境测试保证。

## Consequences

- 开发、测试和打包依赖必须安装在项目虚拟环境中，不得使用系统 Python。
- 开发与打包工具链统一使用 64 位版本，不提供 32 位交付物。
- 便携交付目录必须包含运行所需解释器和依赖，客户机不得承担环境安装步骤。
- 不得使用仅支持 Windows 10/11 的现代 Python 或 GUI 技术栈。
- UI 工具包、Excel 库和打包工具仍需在后续决策中分别确认。
