// 检查是否能打开 autowxwork 协议
function checkAutoWxwork(params) {
  return new Promise((resolve) => {
    const iframe = document.createElement('iframe')
    iframe.style.display = 'none'
    document.body.appendChild(iframe)

    let resolved = false
    const { group, text, image } = params || {}
    
    // 构建检测用的URL，直接带上参数
    let testUrl = 'autowxwork://group=' + group
    if (text)
      testUrl += `&text=${text}`
    if (image) {
      if (typeof image === 'string')
        testUrl += `&image=${image}`
      else if (Array.isArray(image))
        testUrl += `&image=${image.map(item => item.name).join(' ')}`
    }

    // 设置超时检测
    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true
        document.body.removeChild(iframe)
        resolve(false)
      }
    }, 1000)

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

    // 尝试打开带参数的协议
    iframe.src = testUrl
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

  try {
    const canOpen = await checkAutoWxwork(params)  // 传入参数

    if (canOpen) {
      // 如果检测成功，不需要再次调用，因为检测时已经发送了参数
      return true
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
