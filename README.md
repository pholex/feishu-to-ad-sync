# 飞书用户同步到 AD 域

自动将飞书通讯录用户和部门结构同步到 Windows Active Directory 域控制器。

## 功能特性

- 自动同步飞书部门结构到 AD 域 OU
- 创建新用户账号并生成随机密码
- 更新现有用户信息（邮箱、手机号、部门等）
- 检测并处理未匹配用户（禁用并移动到离职 OU）
- 支持拼音重名处理（按工号排序）
- 通过 Union ID 匹配用户（支持改名）
- 发送密码通知邮件给新建用户
- Dry-run 模式预览同步计划

## 使用步骤

### 1. 配置域控制器（首次使用）

#### 安装 OpenSSH 服务器

1. 打开"设置" → "应用" → "可选功能"
2. 点击"添加功能"，搜索并安装"OpenSSH 服务器"
3. 打开"服务"（services.msc），找到"OpenSSH SSH Server"
4. 右键 → 属性 → 启动类型改为"自动" → 启动服务

#### 配置权限

确保登录用户具有 Active Directory 管理权限（Domain Admins 组成员或等效权限）

#### 验证配置

```bash
ssh Administrator@<域控制器IP>
```

成功连接即可使用。

### 2. 创建飞书应用

1. 访问 [飞书开放平台](https://open.feishu.cn/) 创建企业自建应用
2. 获取 App ID 和 App Secret
3. 添加以下权限（应用权限-租户级别）：

```json
{
  "scopes": {
    "tenant": [
      "bitable:app",
      "contact:contact.base:readonly",
      "contact:department.base:readonly",
      "contact:department.organize:readonly",
      "contact:user.base:readonly",
      "contact:user.department:readonly",
      "contact:user.email:readonly",
      "contact:user.employee:readonly",
      "contact:user.employee_id:readonly",
      "contact:user.gender:readonly",
      "contact:user.phone:readonly",
      "directory:department.count:read",
      "directory:department.status:read",
      "directory:department:list",
      "tenant:tenant:readonly"
    ],
    "user": []
  }
}
```

注：`tenant:tenant:readonly` 用于获取企业名称

### 3. 配置环境变量

```bash
# 将 .env.example 文件重命名为 .env
# 编辑 .env 文件，填入飞书应用信息、域控制器配置、邮件配置
```

### 4. 安装依赖并运行

```bash
pip install -r requirements.txt

# 预览模式（推荐首次使用）
python3 sync_to_ad.py --dry-run

# 正式同步
python3 sync_to_ad.py
```

## 同步流程

1. 获取飞书数据（用户和部门）
2. 同步部门结构到 AD 域 OU
3. 对比飞书和 AD 用户，生成同步计划
4. 创建新用户并生成随机密码
5. 更新现有用户信息
6. 发送密码通知邮件

## 字段映射

| CSV字段名 | AD字段 | 飞书API字段 | 说明 |
|----------|--------|------------|------|
| name | DisplayName | name | 显示名称 |
| pinyin | SamAccountName | - | 登录账号（代码生成） |
| enterprise_email | EmailAddress | enterprise_email | 邮箱地址 |
| mobile | Mobile | mobile | 手机号码 |
| employee_no | EmployeeID | employee_no | 员工工号 |
| union_id | EmployeeNumber | union_id | 飞书唯一标识 |
| uuid | info | - | 基于邮箱的固定UUID（代码生成） |

## 注意事项

- 首次使用建议先运行 `--dry-run` 模式查看同步计划
- 输出文件保存在 `./output/` 目录
- 支持青龙面板定时任务，同步完成后会自动发送系统通知
