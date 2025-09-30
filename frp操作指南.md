# frp内网穿透操作指南

## 🎯 根据您的frp配置界面操作

### 第一步：在frp配置界面中设置

根据您提供的图片，请按以下步骤操作：

1. **代理类型**: 保持选择 `http`
2. **代理名称**: 可以保持 `aki` 或改为 `excel-transpose`
3. **内网地址**: 保持 `127.0.0.1`
4. **内网端口**: ⚠️ **重要** - 将 `3000` 修改为 `8080`
5. **子域名**: 可以留空或输入 `aki-excel`
6. **自定义域名**: 输入 `aki-excel.liuchenglu.com`
7. 点击 **"生成"** 按钮

### 第二步：启动Flask应用

在终端中运行：
```bash
# 激活虚拟环境
source venv/bin/activate

# 启动Flask应用
python app.py
```

### 第三步：启动frp客户端

在另一个终端中运行：
```bash
# 进入frp目录
cd frp_0.52.3_darwin_amd64

# 使用我们配置的frpc.ini文件
./frpc -c ../frpc.ini
```

### 第四步：获取公网地址

frp启动成功后，您会看到类似这样的输出：
```
[aki] proxy added success
[aki] start proxy success
```

然后您可以通过以下地址访问：
- **公网地址**: `https://aki-excel.liuchenglu.com`
- **本地地址**: `http://localhost:8080`

## 🔧 如果使用我们预配置的frpc.ini

我们已经为您创建了 `frpc.ini` 配置文件，内容如下：

```ini
[common]
server_addr = frp1.chuantou.org
server_port = 7000

[aki]
type = http
local_ip = 127.0.0.1
local_port = 8080
subdomain = excel-transpose
```

## 📋 完整操作步骤

### 1. 启动Flask应用
```bash
# 终端1
cd "/Users/aki/Documents/AI相关/cursor AI代码练习/转置-思迈特对外报表"
source venv/bin/activate
python app.py
```

### 2. 启动frp客户端
```bash
# 终端2
cd "/Users/aki/Documents/AI相关/cursor AI代码练习/转置-思迈特对外报表/frp_0.52.3_darwin_amd64"
./frpc -c ../frpc.ini
```

### 3. 访问应用
- 本地访问: http://localhost:8080
- 公网访问: https://aki-excel.liuchenglu.com

## ⚠️ 注意事项

1. **端口配置**: 确保内网端口设置为 `8080`，不是 `3000`
2. **Flask应用**: 必须先启动Flask应用，再启动frp客户端
3. **网络连接**: 确保网络连接正常
4. **防火墙**: 确保8080端口没有被防火墙阻止

## 🛠️ 故障排除

### 如果连接失败：
1. 检查Flask应用是否正常运行在8080端口
2. 检查frp配置中的端口是否正确
3. 检查网络连接
4. 查看frp客户端的错误日志

### 如果无法访问公网地址：
1. 确认frp客户端显示 "proxy added success"
2. 等待几分钟让DNS解析生效
3. 尝试使用不同的子域名

## 📞 成功标志

当您看到以下输出时，说明配置成功：
```
[aki] proxy added success
[aki] start proxy success
```

然后您就可以通过公网地址分享您的Excel转置处理工具了！
