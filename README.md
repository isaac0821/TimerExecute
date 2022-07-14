# Timer Execute



#### 简介

- `TimerExecute` 是一个用于定时单次或者多次执行python代码的小工具，同时具备执行日志的记录功能
- 代码使用 `.net framework 4.7.2` 框架

#### 安装

- 复制`TimerExecute\bin\Debug\TimerExecute.exe` 可执行文件到任意目录即可直接使用

#### 使用

- 启动`TimerExecute.exe`

  ![1](https://github.com/isaac0821/TimerExecute/blob/main/Images/1.png)

- 点击`Find`按键，找到需要定时或者重复执行的python代码

- 加载后调整需要执行的时间、次数、频率，此时右下角会显示下一次执行python代码的时间，红色表示尚未确认执行

![2](https://github.com/isaac0821/TimerExecute/blob/main/Images/2.png)

- 点击`Start`按钮，下一次执行时间由红色变为绿色，所指定的python文件将于该时刻开始运行

![3](https://github.com/isaac0821/TimerExecute/blob/main/Images/3.png)

- 随时点击`Now`按钮，可立即执行一次python文件，不影响剩余执行次数
- 每次执行过python文件后，“上一次执行时间”将会被更新，若更新后为绿色，表示执行成功，红色表示python代码运行失败。

![4](https://github.com/isaac0821/TimerExecute/blob/main/Images/4.png)

- 点击`Stop`按钮将结束计时执行，但不会清空设置信息。点击`Reset`按钮将结束计时执行并重置所有输入信息。
- 点击`Log`按钮可以查看执行的情况及python代码的输出信息

![5](https://github.com/isaac0821/TimerExecute/blob/main/Images/5.png)

- 关闭`TimerExecute.exe`后日志文件将自动追加保存到同目录下的`log.txt`中
