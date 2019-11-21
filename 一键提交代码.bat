@echo off
echo 开始上传代码，请输入标题。
echo 标题示例：添加xx功能，修复xx错误...
set/p title=请输入合法的标题：
git add .
git commit -m "%title%"
git push
echo 若上传失败，请执行目录下的“一键同步代码.bat”。
pause