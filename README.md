# batch-email-sender

> windows 10环境演示，通过Excel批量发送邮件的python脚本，支持轮询多个邮件发送、定时发送。

# 教程

1. 安装python环境
2. 运行下面命令下载相关库

   ```python
    pip install pandas openpyxl schedule
   ```
3. 双击运行脚本 `start_email_sender.py`，同层目录会生成文件夹，里面为日志等文件。![2.png](https://blog.zuke.chat/upload/2.png)
4. 填写相关参数![1.png](https://blog.zuke.chat/upload/1.png)
5. excel文件格式说明：

   1. 发件人表：![3.png](https://blog.zuke.chat/upload/3.png)
   2. 收件人表：![4.png](https://blog.zuke.chat/upload/4.png)
6. 避免被163邮箱拦截的建议

   - 减小批量发送规模：每批尝试控制在5封以内
   - 增加发送间隔：批次间隔建议设为180-300秒（3-5分钟）
   - 使用企业邮箱：如果可能，考虑使用163企业邮箱，它的发信限制更宽松
   - 在邮件内容中避免敏感词：某些营销词汇、金融词汇可能触发垃圾邮件过滤
   - 考虑使用专业的邮件营销服务：如果是商业用途，可以考虑使用专业的ESP (Email Service Provider)
7. Github仓库：[batch-email-sender](https://github.com/zu-ke/batch-email-sender)
8. 163邮箱获取SMTP密钥地址：[https://mail.163.com](https://mail.163.com/)
3. **推广：[企业级AI大模型中转站](https://zuke.chat/) ，全球主流大模型稳定低价高并发，感谢老板支持！**

    ![5.png](https://blog.zuke.chat/upload/5.png)