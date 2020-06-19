#### tcp 自连接

shell脚本

```shell
#shell脚本
sysctl -A | grep rane
```

``````shell
netstat 常见参数
    -a (all)显示所有选项，默认不显示LISTEN相关
    -t (tcp)仅显示tcp相关选项
    -u (udp)仅显示udp相关选项
    -n 拒绝显示别名，能显示数字的全部转化成数字。
    -l 仅列出有在 Listen (监听) 的服務状态

    -p 显示建立相关链接的程序名
    -r 显示路由信息，路由表
    -e 显示扩展信息，例如uid等
    -s 按各个协议进行统计
    -c 每隔一个固定时间，执行该netstat命令。
``````

自连接原因：

TCP在发起一个链接的时候会按照计数器的方式选择一个端口号，然后向服务器或者客户端发一个SYN的请求，

在选好一个端口的时候，已经在内核中出现了，

自连接只会在本机出现

解决办法：

​	网络库中会进行很好的判断，getLocalAddress == getRemoteAddress，则出现自连接

#### 计时

 