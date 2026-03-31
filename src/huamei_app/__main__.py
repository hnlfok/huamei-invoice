import toga

class HuameiApp(toga.App):
    def startup(self):
        self.main_window = toga.MainWindow(title=self.formal_name, size=(400, 500))
        box = toga.Box()
        label = toga.Label("华美物流发货单生成器", style=togaStyles.title)
        hint = toga.Label("请使用电脑访问 http://localhost:5000\n或手机访问同一局域网的Web服务",
                          style=togaStyles.hint)
        box.add(label)
        box.add(hint)
        self.main_window.content = box
        self.main_window.show()

togaStyles = type("Styles", (), {
    "title": {"font_size": 20, "padding": 20, "text_align": "center"},
    "hint": {"font_size": 14, "padding": 10, "text_align": "center", "color": "#666"}
})()
