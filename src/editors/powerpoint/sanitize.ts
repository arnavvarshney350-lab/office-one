/** Remove anything that could execute script from stored rich text. */
export function safeHtml(html: string): string {
  const d = new DOMParser().parseFromString(html, 'text/html')
  d.querySelectorAll('script,style,iframe,object,embed,link,meta,form').forEach(n => n.remove())
  d.body.querySelectorAll('*').forEach(el => {
    // copy first: removing an attribute mutates the live NamedNodeMap we are iterating
    // oxlint-disable-next-line unicorn/no-useless-spread
    for (const a of [...el.attributes]) if (/^on/i.test(a.name) || (/^(href|src)$/i.test(a.name) && /^\s*javascript:/i.test(a.value))) el.removeAttribute(a.name)
  })
  return d.body.innerHTML
}

