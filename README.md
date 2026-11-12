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
2. 同步部门结构到 AD 域 OU（以 OU 的 Description 存飞书部门ID）
   - 部门改名/调整上级时移动 OU，临时解除"防止意外删除"保护，移动后恢复
   - 目标位置已有同名 OU 时，并入原 OU 后再移动；部门ID冲突按本次飞书数据判定保留哪个。
     遇到 GPO 链接、计算机/组等非用户对象、无法判定的ID冲突时不自动处理，列出原因
3. 对比飞书和 AD 用户，生成同步计划
4. 创建新用户并生成随机密码
5. 更新现有用户信息（含移动到对应部门 OU）
6. 发送密码通知邮件
7. 清理飞书中已不存在的部门 OU（按部门ID判定，见 `powershell/delete_ou.ps1` 头部注释）
   - 只删除空的残留 OU；含用户等对象或链接了 GPO 的不删，列出原因
   - 位置与飞书部门一致但部门ID不对的 OU，修正为飞书部门ID
   - 残留数超过现有部门 OU 的 20% 时整体跳过（疑似飞书数据不完整）

## 回归测试

PowerShell 脚本的核心逻辑有回归测试，用内存模拟的 AD（`tests/MockAD.ps1`）运行，不需要域控，macOS/Linux/Windows 上装有 PowerShell 7（`pwsh`）即可：

```bash
pwsh -NoProfile -File tests/run.ps1                  # 运行全部测试
TEST_VERBOSE=1 pwsh -NoProfile -File tests/run.ps1   # 同时打印被测脚本的日志
```

| 测试文件 | 覆盖内容 |
|---------|---------|
| `tests/create_ou.Tests.ps1` | 部门 OU 移动时与同名 OU 的合并、停止条件、部门ID冲突按飞书判定、合并中途失败 |
| `tests/create_ou.Main.Tests.ps1` | 整体运行部门同步：改名/移动遇到同名 OU、中间位置被占用、部门ID重复 |
| `tests/delete_ou.Tests.ps1` | 残留 OU 判定与删除、部门ID修正、熔断 |

修改 `powershell/create_ou.ps1` 或 `powershell/delete_ou.ps1` 后请先跑一遍；模拟 AD 只覆盖脚本用到的 cmdlet 行为，上线前仍建议在域控上用 `--dry-run` 预览。

## 字段映射

| CSV字段名 | AD字段 | 飞书API字段 | 说明 |
|----------|--------|------------|------|
| name | DisplayName | name | 显示名称 |
| pinyin | SamAccountName | - | 登录账号（代码生成） |
| enterprise_email | EmailAddress | enterprise_email | 邮箱地址 |
| mobile | Mobile | mobile | 手机号码 |
| employee_no | EmployeeID | employee_no | 员工工号 |
| job_title | Title | job_title | 职务 |
| union_id | EmployeeNumber | union_id | 飞书唯一标识 |
| uuid | info | - | 基于邮箱的固定UUID（代码生成） |

## 注意事项

- 首次使用建议先运行 `--dry-run` 模式查看同步计划
- 输出文件保存在 `./output/` 目录
- 支持青龙面板定时任务，同步完成后会自动发送系统通知
