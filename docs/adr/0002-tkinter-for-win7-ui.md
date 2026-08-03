---
status: accepted
---

# 使用 Tkinter/ttk 构建操作界面

应用使用 Python 3.8 自带的 `Tkinter/ttk` 构建单窗口多标签页 UI。项目是受控流水线中的小型工具，界面只需稳定完成文件和工作表选择、执行及结果提示；选择 Tkinter 可以减少 Win7 便携打包所需的第三方运行库和插件，而 PySide2/Qt5 的体积、部署复杂度与维护成本对本项目没有足够收益。

## Considered Options

- PySide2/Qt5：控件和视觉能力更强，但属于旧版 Qt 技术栈，便携包更大且插件部署更复杂。
- Tkinter/ttk：视觉较传统，但 Python 3.8 自带，足以完成当前多标签页交互，依赖和故障面更小。

## Consequences

- UI 不使用 PySide6、PySide2、Qt Quick 或 Qt Widgets。
- 视觉验收以布局清晰、文字可读和操作正确为主，不建设复杂主题与动画。
- 长时间文件操作不得阻塞窗口刷新；执行状态和最终结果必须在 UI 中可见。
