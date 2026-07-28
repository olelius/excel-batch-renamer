"""桌面应用启动入口。"""

import logging
import sys

from excel_batch_renamer.ui.main_window import MainWindow


def configure_logging() -> None:
    """把运行步骤和异常输出到测试版控制台。"""

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    )


def main() -> None:
    """创建窗口并进入 Tkinter 事件循环。"""

    configure_logging()
    window = MainWindow()
    if "--smoke-test" in sys.argv:
        window.withdraw()
        window.update_idletasks()
        print("ExcelBatchRenamer portable runtime smoke test passed.")
        window.destroy()
        return
    window.mainloop()


if __name__ == "__main__":
    main()
