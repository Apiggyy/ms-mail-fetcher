function fallbackCopyText(text) {
    const textarea = document.createElement('textarea')
    textarea.value = text
    textarea.setAttribute('readonly', '')
    textarea.style.position = 'fixed'
    textarea.style.top = '-9999px'
    textarea.style.left = '-9999px'

    document.body.appendChild(textarea)
    textarea.focus()
    textarea.select()
    textarea.setSelectionRange(0, textarea.value.length)

    const copied = document.execCommand('copy')
    document.body.removeChild(textarea)

    if (!copied) {
        throw new Error('Fallback copy failed')
    }
}

export async function copyText(text) {
    if (!text) return

    if (window.isSecureContext && navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(text)
        return
    }

    fallbackCopyText(text)
}
