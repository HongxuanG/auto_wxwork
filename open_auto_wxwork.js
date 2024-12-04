// 检查是否能打开 autowxwork 协议
function checkAutoWxwork() {
  return new Promise((resolve) => {
    const iframe = document.createElement('iframe')
    iframe.style.display = 'none'
    document.body.appendChild(iframe)

    let resolved = false

    // 设置超时检测
    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true
        document.body.removeChild(iframe)
        resolve(false)
      }
    }, 1000)

    // 监听 iframe 的错误事件
    iframe.onerror = () => {
      if (!resolved) {
        resolved = true
        clearTimeout(timer)
        document.body.removeChild(iframe)
        resolve(false)
      }
    }

    // 如果能打开协议，window.onblur 会被触发
    const handleBlur = () => {
      if (!resolved) {
        resolved = true
        clearTimeout(timer)
        window.removeEventListener('blur', handleBlur)
        document.body.removeChild(iframe)
        resolve(true)
      }
    }

    window.addEventListener('blur', handleBlur)

    // 尝试打开 autowxwork 协议
    iframe.src = 'autowxwork://'
  })
}

// 创建提示div
function createTipDiv() {
  // 检查是否已存在提示div
  const existingTip = document.getElementById('wxwork-tip')
  if (existingTip)
    return existingTip

  const tipDiv = document.createElement('div')
  tipDiv.id = 'wxwork-tip'
  tipDiv.style.cssText = `
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    background: #fff;
    padding: 20px;
    border-radius: 4px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.1);
    z-index: 9999;
    text-align: center;
  `

  tipDiv.innerHTML = `
    未检测到企业微信自动化工具，请先
    <a 
      href="/企业微信自动化工具安装程序.exe" 
      style="color: #1890ff; text-decoration: underline;"
      download
    >
      安装
    </a>
  `

  // 添加关闭按钮
  const closeBtn = document.createElement('div')
  closeBtn.style.cssText = `
    position: absolute;
    right: 10px;
    top: 10px;
    cursor: pointer;
    color: #999;
  `
  closeBtn.innerHTML = '✕'
  closeBtn.onclick = () => document.body.removeChild(tipDiv)

  tipDiv.appendChild(closeBtn)
  return tipDiv
}

// 主函数
export async function openAutoWxwork(params) {
  const { text, group, image } = params
  let str = ''

  if (text)
    str += `&text=${text}`
  if (image) {
    if (typeof image === 'string')
      str += `&image=${image}`
    else if (Array.isArray(image))
      str += `&image=${image.map(item => item.name).join(' ')}`
  }

  try {
    const canOpen = await checkAutoWxwork()

    if (canOpen) {
      // 如果能打开协议，直接打开
      window.location.href = `autowxwork://group=${group}${str}`
    }
    else {
      // 如果不能打开，显示提示div
      const tipDiv = createTipDiv()
      document.body.appendChild(tipDiv)
    }
  }
  catch (error) {
    console.error('打开企业微信自动化工具失败:', error)
  }
}

// 导出检查函数，方便其他地方使用
export function checkWxworkInstalled() {
  return checkAutoWxwork()
}
