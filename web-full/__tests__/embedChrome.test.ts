// @vitest-environment jsdom
// The editor's own toasts and AI entries in the embed: the host reports the real save outcome, so the English
// "Saved" toast (the app's ToastHost, .app-toast.success) stays hidden; error toasts still show. The floating
// "Ask AI" button is hidden too.
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

const hostCss = readFileSync(resolve(__dirname, '../host.css'), 'utf8')

afterEach(() => {
  document.head.innerHTML = ''
  document.body.innerHTML = ''
  document.documentElement.classList.remove('vound-embedded', 'vound-review', 'vound-fields-only', 'vound-view')
})

function element(className: string) {
  const style = document.createElement('style')
  style.textContent = hostCss
  document.head.append(style)
  const node = document.createElement('div')
  node.className = className
  document.body.append(node)
  return node
}

describe('editor chrome in the embed', () => {
  it('hides what reshapes the document only while locked to a document with fields', () => {
    const handle = element('doc-table-resize')
    const field = document.createElement('span')
    field.className = 'vound-slot-field'
    document.body.append(field)
    expect(getComputedStyle(handle).display).not.toBe('none')

    document.documentElement.classList.add('vound-review', 'vound-fields-only')
    expect(getComputedStyle(handle).display).toBe('none')

    field.remove()
    expect(getComputedStyle(handle).display).not.toBe('none')
  })

  it('hides the success toast when embedded and nowhere else', () => {
    const saved = element('app-toast success show')
    expect(getComputedStyle(saved).display).not.toBe('none')
    document.documentElement.classList.add('vound-embedded')
    expect(getComputedStyle(saved).display).toBe('none')
  })

  it('keeps error toasts', () => {
    document.documentElement.classList.add('vound-embedded')
    expect(getComputedStyle(element('app-toast error show')).display).not.toBe('none')
  })

  it('hides the Ask AI button when embedded and nowhere else', () => {
    const ask = element('ai-ask-trigger')
    expect(getComputedStyle(ask).display).not.toBe('none')
    document.documentElement.classList.add('vound-embedded')
    expect(getComputedStyle(ask).display).toBe('none')
  })

  it('shows only the pages in view mode, and they take no pointer input', () => {
    const chrome = ['ribbon', 'ribbon-tabs', 'status-bar'].map(element)
    const page = element('doc-page')
    expect(chrome.map((node) => getComputedStyle(node).display)).not.toContain('none')
    expect(getComputedStyle(page).pointerEvents).not.toBe('none')

    document.documentElement.classList.add('vound-view')
    expect(chrome.map((node) => getComputedStyle(node).display)).toEqual(['none', 'none', 'none'])
    expect(getComputedStyle(page).pointerEvents).toBe('none')
  })
})
