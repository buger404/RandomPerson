@echo off
echo ��ʼ�ϴ����룬��������⡣
echo ����ʾ�������xx���ܣ��޸�xx����...
set/p title=������Ϸ��ı��⣺
git add .
git commit -m "%title%"
git push
echo ���ϴ�ʧ�ܣ���ִ��Ŀ¼�µġ�һ��ͬ������.bat����
pause