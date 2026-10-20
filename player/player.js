/*
 * Advanced Exporter — slide player.
 *
 * The slides are pictures PowerPoint drew. This puts back what a picture
 * can't carry: keys, clicks and swipes to move through them, the links and
 * videos over the right spots, the words for screen readers and Find, and a
 * presenter view with the notes.
 *
 * Plain ES2017, no dependencies, no build step: the add-in writes this file
 * into the export as it is, and the export has to work opened straight from
 * disk (file://). That is why the presenter window is kept in step with
 * postMessage between the two window references — BroadcastChannel can't be
 * relied on for file:// pages, so it is only used as an extra.
 *
 * The deck arrives as JSON in <script type="application/json" id="deck-data">.
 * Its shape is PlayerDeck in src/formats/web.ts.
 */
(function () {
  'use strict'

  var NS = 'advanced-exporter-player'
  var DEFAULT_MS = 500
  var IDLE_MS = 2500
  var REDUCED = '(prefers-reduced-motion: reduce)'
  var SAFE_HREF = /^(https?:|mailto:|tel:)/i
  // Which way the slides travel: [x, y] as fractions of the slide.
  var VEC = { left: [-1, 0], right: [1, 0], up: [0, -1], down: [0, 1] }
  var OPPOSITE = { left: 'right', right: 'left', up: 'down', down: 'up' }

  var api = { version: 1, current: null, start: start }
  if (window.AdvancedPlayer && window.AdvancedPlayer.current) {
    try { window.AdvancedPlayer.current.destroy() } catch (e) { /* an older copy */ }
  }
  window.AdvancedPlayer = api

  // ---------------------------------------------------------------- helpers

  /** Build an element. `html` is only ever our own static markup (icons). */
  function h(tag, props, kids) {
    var e = document.createElement(tag)
    if (props) {
      for (var k in props) {
        if (!Object.prototype.hasOwnProperty.call(props, k)) continue
        var v = props[k]
        if (v == null || v === false) continue
        if (k === 'className') e.className = v
        else if (k === 'text') e.textContent = v
        else if (k === 'html') e.innerHTML = v
        else e.setAttribute(k, v === true ? '' : String(v))
      }
    }
    if (kids) {
      for (var i = 0; i < kids.length; i++) {
        var c = kids[i]
        if (c) e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c)
      }
    }
    return e
  }

  function Listeners() { this.list = [] }
  Listeners.prototype.on = function (target, type, fn, opts) {
    target.addEventListener(type, fn, opts || false)
    this.list.push([target, type, fn, opts || false])
  }
  Listeners.prototype.off = function () {
    this.list.forEach(function (l) { l[0].removeEventListener(l[1], l[2], l[3]) })
    this.list = []
  }

  function mq(query) {
    try { return !!(window.matchMedia && window.matchMedia(query).matches) } catch (e) { return false }
  }

  function now() { return Date.now() }

  function closest(t, sel) {
    return t && t.nodeType === 1 && t.closest ? t.closest(sel) : null
  }

  function isEditable(t) {
    if (!t || t.nodeType !== 1) return false
    if (t.isContentEditable) return true
    var tag = t.tagName
    if (tag === 'TEXTAREA' || tag === 'SELECT') return true
    if (tag === 'INPUT') return !/^(button|submit|reset|checkbox|radio|range|color|file|image)$/i.test(t.type || '')
    return false
  }

  /** A '#id' source means "the src of that element" — a single-file export
   * keeps each picture once, in its no-script fallback, and points here. */
  function resolveSrc(src) {
    src = src == null ? '' : String(src)
    if (src.charAt(0) === '#') {
      var n = document.getElementById(src.slice(1))
      return n ? n.getAttribute('src') || '' : ''
    }
    return src
  }

  function store(key, value) {
    try {
      if (value === undefined) return window.localStorage.getItem(NS + ':' + key)
      window.localStorage.setItem(NS + ':' + key, String(value))
    } catch (e) { /* private mode, file:// in some browsers */ }
    return null
  }

  function slideLabel(s, i, n) {
    return 'Slide ' + (i + 1) + ' of ' + n + (s.title ? ': ' + s.title : '')
  }

  var ICONS = {
    prev: '<path d="M15 18l-6-6 6-6"/>',
    next: '<path d="M9 18l6-6-6-6"/>',
    grid: '<rect x="4" y="4" width="6" height="6" rx="1"/><rect x="14" y="4" width="6" height="6" rx="1"/><rect x="4" y="14" width="6" height="6" rx="1"/><rect x="14" y="14" width="6" height="6" rx="1"/>',
    full: '<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>',
    unfull: '<path d="M9 4v5H4M15 4v5h5M9 20v-5H4M15 20v-5h5"/>',
    presenter: '<rect x="3" y="4" width="18" height="12" rx="1"/><path d="M8 20h8M12 16v4"/>',
    help: '<circle cx="12" cy="12" r="9"/><path d="M9.6 9.4a2.5 2.5 0 1 1 3.4 2.4c-.6.3-1 .8-1 1.4V14M12 17.5v.01"/>',
    pause: '<path d="M9 5v14M15 5v14"/>',
    play: '<path d="M7 5l12 7-12 7z"/>',
    close: '<path d="M6 6l12 12M18 6L6 18"/>',
    reset: '<path d="M4 12a8 8 0 1 0 2.3-5.7M4 4v4h4"/>',
    smaller: '<path d="M5 12h14"/>',
    bigger: '<path d="M5 12h14M12 5v14"/>',
    blank: '<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M3 5l18 14"/>',
  }

  function icon(name) {
    return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">' + ICONS[name] + '</svg>'
  }

  function button(L, name, label, onClick, extraClass) {
    var b = h('button', { type: 'button', className: 'ap-btn' + (extraClass ? ' ' + extraClass : ''), 'aria-label': label, title: label, html: icon(name) })
    L.on(b, 'click', function (e) {
      e.stopPropagation()
      if (b.getAttribute('aria-disabled') === 'true') return
      onClick()
    })
    return b
  }

  function setButton(b, name, label) {
    if (!b) return
    b.innerHTML = icon(name)
    b.setAttribute('aria-label', label)
    b.title = label
  }

  // ---------------------------------------------------------------- data

  function normalise(raw) {
    var o = raw.options || {}
    var slides = Array.isArray(raw.slides) ? raw.slides.filter(function (s) { return s && typeof s === 'object' }) : []
    return {
      title: String(raw.title || ''),
      aspect: Number(raw.aspect) > 0 ? Number(raw.aspect) : 16 / 9,
      background: typeof raw.background === 'string' ? raw.background : '#000000',
      slides: slides,
      options: {
        controls: o.controls !== false,
        progress: o.progress !== false,
        numbers: o.numbers !== false,
        loop: !!o.loop,
        auto: o.auto === 'deck' || o.auto === 'every' ? o.auto : 'off',
        autoSeconds: Number(o.autoSeconds) > 0 ? Number(o.autoSeconds) : 5,
        transitions: /^(deck|fade|slide|none)$/.test(o.transitions) ? o.transitions : 'deck',
        transitionMs: Number(o.transitionMs) > 0 ? Math.min(Number(o.transitionMs), 10000) : 0,
        theme: o.theme === 'light' ? 'light' : 'dark',
        presenter: o.presenter !== false,
        skipHidden: o.skipHidden !== false,
      },
    }
  }

  /** Moving through the deck: hidden slides are stepped over (PowerPoint
   * does the same) but can still be reached by a link or a number. */
  function makeNav(slides, opts) {
    var N = slides.length
    var skip = opts.skipHidden && slides.some(function (s) { return !s.hidden })
    function visible(j) { return !(skip && slides[j].hidden) }
    return {
      N: N,
      step: function (i, d) {
        var j = i
        for (var k = 0; k < N; k++) {
          j += d
          if (j >= N || j < 0) {
            if (!opts.loop) return -1
            j = j < 0 ? N - 1 : 0
          }
          if (visible(j)) return j
        }
        return -1
      },
      first: function () { for (var j = 0; j < N; j++) if (visible(j)) return j; return 0 },
      last: function () { for (var j = N - 1; j >= 0; j--) if (visible(j)) return j; return N - 1 },
    }
  }

  /** What a key does, wherever it is pressed. */
  function keyAction(e) {
    switch (e.key) {
      case 'ArrowRight': case 'Right': case 'ArrowDown': case 'Down': case 'PageDown': case 'n': case 'N':
        return 'next'
      case ' ': case 'Spacebar':
        return e.shiftKey ? 'prev' : 'next'
      case 'Enter':
        return 'next'
      case 'ArrowLeft': case 'Left': case 'ArrowUp': case 'Up': case 'PageUp': case 'Backspace': case 'p': case 'P':
        return 'prev'
      case 'Home': return 'first'
      case 'End': return 'last'
      case 'f': case 'F': return 'full'
      case 'b': case 'B': case '.': return 'black'
      case 'w': case 'W': case ',': return 'white'
      case 'o': case 'O': return 'overview'
      case 'Escape': case 'Esc': return 'escape'
      case 's': case 'S': return 'presenter'
      case '?': return 'help'
    }
    return null
  }

  /** Typing a slide number, then Enter. Forgets after a few idle seconds. */
  function makeNumberEntry(show) {
    var digits = ''
    var timer = 0
    function set(d) {
      digits = d
      clearTimeout(timer)
      if (digits) timer = setTimeout(function () { set('') }, 4000)
      show(digits)
    }
    return {
      value: function () { return digits },
      push: function (c) { set((digits + c).slice(-4)) },
      back: function () { set(digits.slice(0, -1)) },
      clear: function () { set('') },
      stop: function () { clearTimeout(timer) },
    }
  }

  // ---------------------------------------------------------------- full screen

  function fullscreenElement() {
    return document.fullscreenElement || document.webkitFullscreenElement || null
  }

  function canFullscreen() {
    var d = document.documentElement
    if (!(d.requestFullscreen || d.webkitRequestFullscreen)) return false
    return document.fullscreenEnabled !== false || !!document.webkitFullscreenEnabled
  }

  function toggleFullscreen() {
    try {
      var p
      if (fullscreenElement()) p = (document.exitFullscreen || document.webkitExitFullscreen).call(document)
      else {
        var d = document.documentElement
        p = (d.requestFullscreen || d.webkitRequestFullscreen).call(d)
      }
      if (p && p.catch) p.catch(function () { /* refused: not from a user gesture */ })
    } catch (e) { /* not supported here */ }
  }

  // ---------------------------------------------------------------- window sync

  /** Keeps the audience and presenter windows on the same slide. Every change
   * goes to the other window directly (opener / child reference — works from
   * file://) and to a BroadcastChannel where there is one (a presenter view
   * opened in another tab over http). Remote changes are applied, never
   * re-sent, so nothing loops. */
  function makeSync(key, onState, onHello, L) {
    var id = Math.random().toString(36).slice(2) + now().toString(36)
    var seq = 0
    var seen = {}
    var child = null
    var channel = null
    var heard = 0

    function alive(w) {
      try { return !!w && w !== window && !w.closed } catch (e) { return false }
    }
    function peers() {
      var out = []
      if (alive(child)) out.push(child)
      if (alive(window.opener) && window.opener !== child) out.push(window.opener)
      return out
    }
    function post(type, body) {
      var msg = { ns: NS, key: key, from: id, seq: ++seq, type: type, body: body || null }
      peers().forEach(function (w) {
        try { w.postMessage(msg, '*') } catch (e) { /* gone */ }
      })
      if (channel) {
        try { channel.postMessage(msg) } catch (e) { /* closed */ }
      }
    }
    function receive(msg, source) {
      if (!msg || typeof msg !== 'object' || msg.ns !== NS || msg.key !== key || msg.from === id) return
      if (typeof msg.seq !== 'number' || (seen[msg.from] || 0) >= msg.seq) return
      seen[msg.from] = msg.seq
      heard = now()
      // A presenter that outlived a reload of this page finds us again here.
      if (source && !alive(child) && source !== window.opener) child = source
      if (msg.type === 'hello') onHello()
      else if (msg.type === 'state' && msg.body) onState(msg.body)
    }
    L.on(window, 'message', function (e) { receive(e.data, e.source) })
    try {
      if (typeof BroadcastChannel === 'function') {
        channel = new BroadcastChannel(NS + ':' + key)
        channel.onmessage = function (e) { receive(e.data, null) }
      }
    } catch (e) { channel = null }

    return {
      post: post,
      setChild: function (w) { child = w },
      child: function () { return alive(child) ? child : null },
      // A window reference, or word from another window in the last minute.
      linked: function () { return peers().length > 0 || now() - heard < 60000 },
      close: function () {
        if (channel) {
          try { channel.close() } catch (e) { /* already */ }
        }
      },
    }
  }

  function deckKey(data) {
    var first = data.slides[0] ? String(data.slides[0].src).slice(0, 64) : ''
    return data.title + '|' + data.slides.length + '|' + first
  }

  // ---------------------------------------------------------------- start

  function start() {
    if (api.current) api.current.destroy()
    var node = document.getElementById('deck-data')
    if (!node) return null
    var data
    try {
      data = normalise(JSON.parse(node.textContent || '{}'))
    } catch (e) {
      if (window.console) console.error('Slide player: the deck data could not be read.', e)
      return null
    }
    if (!data.slides.length) return null
    var host = document.getElementById('player')
    if (!host) {
      host = h('div', { id: 'player' })
      document.body.appendChild(host)
    }
    var presenter = /^#presenter\b/.test(location.hash) && data.options.presenter
    api.current = presenter ? createPresenter(data, host) : createAudience(data, host)
    return api.current
  }

  // ---------------------------------------------------------------- audience

  function createAudience(data, host) {
    var L = new Listeners()
    var slides = data.slides
    var opts = data.options
    var nav = makeNav(slides, opts)
    var N = nav.N
    var index = -1
    var blank = ''
    var overlay = ''
    var view = null
    var running = []
    var blobs = {}
    var preloaded = {}
    var auto = { timer: 0, deadline: 0, left: -1, holds: {} }
    var idleTimer = 0
    var toastTimer = 0
    var barHover = false
    var touch = null
    var swipedAt = 0
    var thumbs = null
    var lastFocus = null

    document.documentElement.classList.add('ap-on')
    var root = h('div', { className: 'ap ap-audience', 'data-theme': opts.theme })
    root.style.setProperty('--ap-slide-bg', data.background)
    root.style.setProperty('--ap-aspect', String(data.aspect))

    var stage = h('div', { className: 'ap-stage', tabindex: '-1' })
    var frame = h('div', { className: 'ap-frame', role: 'group', 'aria-roledescription': 'slide' })
    stage.appendChild(frame)
    root.appendChild(stage)

    var blankEl = h('div', { className: 'ap-blank', hidden: true })
    root.appendChild(blankEl)

    var progress = null
    if (opts.progress) {
      progress = h('div', { className: 'ap-progress', 'aria-hidden': 'true' }, [h('div', { className: 'ap-progress-bar' })])
      root.appendChild(progress)
    }

    var counter = opts.numbers ? h('span', { className: 'ap-counter' + (opts.controls ? '' : ' ap-alone'), 'aria-hidden': 'true' }) : null
    var btn = {}
    var bar = null
    if (opts.controls) {
      bar = h('div', { className: 'ap-bar', role: 'toolbar', 'aria-label': 'Slide controls' })
      btn.prev = button(L, 'prev', 'Previous slide', prev)
      btn.next = button(L, 'next', 'Next slide', next)
      bar.appendChild(btn.prev)
      if (counter) bar.appendChild(counter)
      bar.appendChild(btn.next)
      var timed = opts.auto === 'every' || (opts.auto === 'deck' && slides.some(function (s) { return typeof s.advanceMs === 'number' }))
      if (timed) {
        btn.auto = button(L, 'pause', 'Pause automatic advance', toggleAutoPause)
        bar.appendChild(btn.auto)
      }
      btn.overview = button(L, 'grid', 'All slides (O)', function () { toggleOverlay('overview') })
      bar.appendChild(btn.overview)
      if (canFullscreen()) {
        btn.full = button(L, 'full', 'Full screen (F)', toggleFullscreen)
        bar.appendChild(btn.full)
      }
      if (opts.presenter) {
        btn.presenter = button(L, 'presenter', 'Presenter view (S)', openPresenter)
        bar.appendChild(btn.presenter)
      }
      btn.help = button(L, 'help', 'Keyboard shortcuts (?)', function () { toggleOverlay('help') })
      bar.appendChild(btn.help)
      root.appendChild(bar)
      L.on(bar, 'mouseenter', function () { barHover = true })
      L.on(bar, 'mouseleave', function () { barHover = false })
    } else if (counter) {
      root.appendChild(counter)
    }

    var gotoEl = h('div', { className: 'ap-goto', hidden: true, 'aria-hidden': 'true' })
    var toastEl = h('div', { className: 'ap-toast', role: 'status', hidden: true })
    var live = h('div', { className: 'ap-sr', 'aria-live': 'polite', 'aria-atomic': 'true' })
    root.appendChild(gotoEl)
    root.appendChild(toastEl)
    root.appendChild(live)

    var overviewEl = h('div', { className: 'ap-overview', role: 'dialog', 'aria-modal': 'true', 'aria-label': 'All slides', hidden: true })
    var helpEl = buildHelp()
    root.appendChild(overviewEl)
    root.appendChild(helpEl)

    var textRegion = buildTextRegion()
    root.appendChild(textRegion)

    host.appendChild(root)

    var numbers = makeNumberEntry(function (d) {
      gotoEl.hidden = !d
      gotoEl.textContent = d ? 'Go to slide ' + d + ' ↵' : ''
    })

    var sync = makeSync(deckKey(data), function (st) {
      if (typeof st.index === 'number') go(st.index, { remote: true })
      if (typeof st.blank === 'string') setBlank(st.blank, true)
    }, function () {
      sync.post('state', state())
    }, L)

    // ---- the text a picture can't give a screen reader or Find

    function linkEl(l, cls) {
      var a = h('a', cls ? { className: cls } : null)
      if (l.href && SAFE_HREF.test(l.href)) {
        a.href = l.href
        a.target = '_blank'
        a.rel = 'noopener noreferrer'
      } else if (typeof l.slide === 'number' && l.slide >= 0 && l.slide < N) {
        a.href = '#/' + (l.slide + 1)
        a.setAttribute('data-slide', String(l.slide))
      } else {
        return null
      }
      if (l.tooltip) a.title = l.tooltip
      return a
    }

    function linkLabel(l) {
      var label = l.label || l.tooltip || (l.href ? l.href : 'Slide ' + (l.slide + 1))
      return l.href ? label + ' (opens in a new tab)' : label
    }

    function buildTextRegion() {
      var region = h('div', { className: 'ap-sr', role: 'region', 'aria-label': 'Text of the slides' })
      slides.forEach(function (s, i) {
        var sec = h('section', { 'data-index': i })
        sec.appendChild(h('h2', { text: 'Slide ' + (i + 1) + (s.title ? ': ' + s.title : '') }))
        String(s.text || '').split('\n').forEach(function (line) {
          if (line.trim()) sec.appendChild(h('p', { text: line }))
        })
        var links = (s.links || []).map(function (l) {
          var a = linkEl(l, null)
          if (!a) return null
          a.textContent = linkLabel(l)
          // Reachable by a screen reader's own cursor; kept out of the Tab
          // order, where it would be an invisible focus stop.
          a.setAttribute('tabindex', '-1')
          return h('li', null, [a])
        }).filter(Boolean)
        if (links.length) sec.appendChild(h('ul', null, links))
        region.appendChild(sec)
      })
      return region
    }

    // ---- one slide on the stage

    function hasBox(b) {
      return b && typeof b.x === 'number' && typeof b.y === 'number' && b.w > 0 && b.h > 0
    }

    function place(e, b) {
      e.style.left = b.x * 100 + '%'
      e.style.top = b.y * 100 + '%'
      e.style.width = b.w * 100 + '%'
      e.style.height = b.h * 100 + '%'
    }

    /** Embedded media in a single-file export is a data: URI; a Blob URL
     * plays and seeks better, so it is turned into one when first needed. */
    function mediaUrl(m, key) {
      var src = resolveSrc(m.src)
      if (!/^data:/i.test(src)) return src
      if (blobs[key]) return blobs[key]
      try {
        var comma = src.indexOf(',')
        var meta = src.slice(5, comma)
        var mime = meta.split(';')[0] || m.type || 'application/octet-stream'
        var body = src.slice(comma + 1)
        var bin = /;base64/i.test(meta) ? atob(body) : decodeURIComponent(body)
        var bytes = new Uint8Array(bin.length)
        for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
        blobs[key] = URL.createObjectURL(new Blob([bytes], { type: mime }))
        return blobs[key]
      } catch (e) {
        return src
      }
    }

    function makeView(i) {
      var s = slides[i]
      var v = h('div', { className: 'ap-slide', 'data-index': i })
      var img = h('img', { className: 'ap-img', alt: s.alt || s.title || 'Slide ' + (i + 1), draggable: 'false', decoding: 'async' })
      img.src = resolveSrc(s.src)
      v.appendChild(img)
      ;(s.links || []).forEach(function (l) {
        if (!hasBox(l)) return
        var a = linkEl(l, 'ap-link')
        if (!a) return
        a.setAttribute('aria-label', linkLabel(l))
        place(a, l)
        v.appendChild(a)
      })
      ;(s.media || []).forEach(function (m, j) {
        if (!hasBox(m) || !m.src) return
        var audio = m.kind === 'audio'
        var box = h('div', { className: 'ap-mediabox' + (audio ? ' is-audio' : '') })
        place(box, m)
        var el = document.createElement(audio ? 'audio' : 'video')
        el.controls = true
        el.preload = 'metadata'
        el.setAttribute('playsinline', '')
        el.setAttribute('aria-label', m.label || (audio ? 'Audio' : 'Video'))
        if (m.poster && !audio) el.poster = resolveSrc(m.poster)
        el.src = mediaUrl(m, i + ':' + j)
        el.addEventListener('play', mediaChanged)
        el.addEventListener('pause', mediaChanged)
        el.addEventListener('ended', mediaChanged)
        box.appendChild(el)
        v.appendChild(box)
      })
      return v
    }

    function mediaIn(v) {
      return v ? Array.prototype.slice.call(v.querySelectorAll('video, audio')) : []
    }

    function pauseMedia(v) {
      mediaIn(v).forEach(function (m) {
        try { m.pause() } catch (e) { /* not playing */ }
      })
    }

    function mediaChanged() {
      var playing = mediaIn(view).some(function (m) { return !m.paused && !m.ended })
      if (playing) hold('media')
      else release('media')
    }

    // ---- transitions

    function transitionFor(from, to, forward) {
      if (from < 0 || opts.transitions === 'none' || mq(REDUCED)) return null
      if (typeof frame.animate !== 'function') return null
      // A transition belongs to the slide it brings in; going back plays the
      // one you came through, reversed.
      var tr = (slides[forward ? to : from] || {}).transition || {}
      var type = opts.transitions === 'deck' ? tr.type : opts.transitions
      if (type !== 'fade' && type !== 'slide') return null
      var ms = opts.transitionMs || tr.ms || DEFAULT_MS
      if (!(ms > 0)) return null
      var dir = opts.transitions === 'deck' && VEC[tr.dir] ? tr.dir : 'left'
      if (!forward) dir = OPPOSITE[dir]
      return { type: type, ms: Math.min(ms, 10000), dir: dir }
    }

    /** Finish whatever transition is still running, at once. */
    function settle() {
      var r = running
      running = []
      r.forEach(function (a) {
        a.onfinish = null
        try { a.cancel() } catch (e) { /* done */ }
      })
      Array.prototype.slice.call(frame.children).forEach(function (c) {
        if (c !== view) frame.removeChild(c)
      })
    }

    function showView(i, from, forward) {
      settle()
      var old = view
      if (old) pauseMedia(old)
      release('media')
      var v = makeView(i)
      frame.appendChild(v)
      view = v
      var t = old ? transitionFor(from, i, forward) : null
      if (!t) {
        if (old && old.parentNode) old.parentNode.removeChild(old)
        return
      }
      var timing = { duration: t.ms, easing: 'ease-in-out', fill: 'both' }
      var anims = []
      try {
        if (t.type === 'fade') {
          anims.push(v.animate([{ opacity: 0 }, { opacity: 1 }], timing))
        } else {
          var d = VEC[t.dir]
          var out = 'translate(' + d[0] * 100 + '%,' + d[1] * 100 + '%)'
          var inn = 'translate(' + -d[0] * 100 + '%,' + -d[1] * 100 + '%)'
          anims.push(old.animate([{ transform: 'translate(0,0)' }, { transform: out }], timing))
          anims.push(v.animate([{ transform: inn }, { transform: 'translate(0,0)' }], timing))
        }
      } catch (e) {
        anims.forEach(function (a) { a.cancel() })
        if (old.parentNode) old.parentNode.removeChild(old)
        return
      }
      running = anims
      anims[anims.length - 1].onfinish = function () {
        if (running === anims) settle()
      }
    }

    // ---- moving

    function state() {
      return { index: index, blank: blank }
    }

    function go(i, how) {
      how = how || {}
      if (typeof i !== 'number' || isNaN(i)) return
      i = Math.max(0, Math.min(N - 1, Math.floor(i)))
      if (blank && !how.remote) setBlank('', true)
      if (i === index) return
      var from = index
      index = i
      var forward = typeof how.forward === 'boolean' ? how.forward : i > from
      showView(i, from, forward)
      update()
      preload(i)
      if (!how.fromHash) writeHash(i, from >= 0 && !how.remote && !how.auto)
      armAuto()
      if (!how.remote) sync.post('state', state())
    }

    function next() {
      if (blank) return setBlank('')
      var j = nav.step(index, 1)
      if (j >= 0) go(j, { forward: true })
    }

    function prev() {
      if (blank) return setBlank('')
      var j = nav.step(index, -1)
      if (j >= 0) go(j, { forward: false })
    }

    function update() {
      var s = slides[index]
      var label = slideLabel(s, index, N)
      live.textContent = label
      frame.setAttribute('aria-label', label)
      if (counter) counter.textContent = index + 1 + ' / ' + N
      if (progress) progress.firstChild.style.width = ((index + 1) / N) * 100 + '%'
      if (btn.prev) btn.prev.setAttribute('aria-disabled', String(nav.step(index, -1) < 0))
      if (btn.next) btn.next.setAttribute('aria-disabled', String(nav.step(index, 1) < 0))
      Array.prototype.forEach.call(textRegion.children, function (sec, j) {
        if (j === index) sec.setAttribute('aria-current', 'true')
        else sec.removeAttribute('aria-current')
      })
      if (thumbs) markThumb()
    }

    function preload(i) {
      var n1 = nav.step(i, 1)
      var n2 = n1 >= 0 ? nav.step(n1, 1) : -1
      ;[n1, n2, nav.step(i, -1)].forEach(function (j) {
        if (j < 0 || preloaded[j]) return
        preloaded[j] = true
        var im = new Image()
        im.decoding = 'async'
        im.src = resolveSrc(slides[j].src)
      })
    }

    // ---- the address bar: #/7 is slide 7

    function hashIndex() {
      var m = /^#\/(\d+)/.exec(location.hash || '')
      if (!m) return -1
      var n = parseInt(m[1], 10) - 1
      return n >= 0 && n < N ? n : -1
    }

    function writeHash(i, push) {
      var hash = '#/' + (i + 1)
      if (location.hash === hash) return
      try {
        history[push ? 'pushState' : 'replaceState'](null, '', hash)
      } catch (e) {
        // Some browsers refuse history entries for file:// pages.
        try { location.replace(hash) } catch (e2) { /* leave the address alone */ }
      }
    }

    function onHash() {
      var i = hashIndex()
      if (i >= 0 && i !== index) go(i, { fromHash: true })
    }

    // ---- black and white screens

    function setBlank(mode, quiet) {
      mode = mode === 'black' || mode === 'white' ? mode : ''
      if (mode === blank) return
      blank = mode
      blankEl.hidden = !mode
      blankEl.setAttribute('data-mode', mode)
      if (mode) hold('blank')
      else release('blank')
      live.textContent = mode ? 'Screen ' + mode + '. Press any key to go back.' : slideLabel(slides[index], index, N)
      if (!quiet) sync.post('state', state())
    }

    // ---- automatic advance

    function autoDelay(i) {
      if (opts.auto === 'every') return opts.autoSeconds * 1000
      // PowerPoint's "After 0s" means straight on; a floor keeps a deck of
      // zeros from spinning.
      var a = slides[i].advanceMs
      if (opts.auto === 'deck') return typeof a === 'number' && a >= 0 ? Math.max(a, 250) : -1
      return -1
    }

    function held() {
      return Object.keys(auto.holds).length > 0
    }

    function armAuto() {
      clearTimeout(auto.timer)
      auto.timer = 0
      auto.left = autoDelay(index)
      runAuto()
    }

    function runAuto() {
      if (auto.left < 0 || held() || auto.timer) return
      auto.deadline = now() + auto.left
      auto.timer = setTimeout(function () {
        auto.timer = 0
        auto.left = -1
        var j = nav.step(index, 1)
        if (j >= 0) go(j, { forward: true, auto: true })
      }, auto.left)
    }

    function hold(reason) {
      if (auto.timer) {
        clearTimeout(auto.timer)
        auto.timer = 0
        auto.left = Math.max(0, auto.deadline - now())
      }
      auto.holds[reason] = true
    }

    function release(reason) {
      if (!auto.holds[reason]) return
      delete auto.holds[reason]
      runAuto()
    }

    function toggleAutoPause() {
      if (auto.holds.user) release('user')
      else hold('user')
      var paused = !!auto.holds.user
      setButton(btn.auto, paused ? 'play' : 'pause', paused ? 'Resume automatic advance' : 'Pause automatic advance')
      toast(paused ? 'Automatic advance paused' : 'Automatic advance resumed')
    }

    // ---- overview and help

    function buildOverview() {
      var head = h('div', { className: 'ap-dialog-head' }, [h('h2', { text: 'All slides' })])
      head.appendChild(button(L, 'close', 'Close (Esc)', closeOverlay, 'ap-close'))
      var grid = h('div', { className: 'ap-grid' })
      thumbs = slides.map(function (s, i) {
        var b = h('button', {
          type: 'button',
          className: 'ap-thumb' + (s.hidden ? ' is-hidden' : ''),
          'data-index': i,
          'aria-label': 'Slide ' + (i + 1) + (s.title ? ': ' + s.title : '') + (s.hidden ? ' (hidden)' : ''),
        })
        // loading before src, or the browser starts fetching straight away.
        var img = h('img', { alt: '', loading: 'lazy', decoding: 'async' })
        img.src = resolveSrc(s.src)
        b.appendChild(h('span', { className: 'ap-thumb-pic' }, [img]))
        b.appendChild(h('span', { className: 'ap-thumb-label' }, [h('b', { text: String(i + 1) }), s.title ? ' ' + s.title : '']))
        L.on(b, 'click', function () {
          closeOverlay()
          go(i)
        })
        grid.appendChild(b)
        return b
      })
      overviewEl.appendChild(head)
      overviewEl.appendChild(grid)
    }

    function markThumb() {
      thumbs.forEach(function (b, j) {
        if (j === index) b.setAttribute('aria-current', 'true')
        else b.removeAttribute('aria-current')
      })
    }

    function buildHelp() {
      var rows = [
        ['Next slide', '→, ↓, Space, Page Down, Enter, N or a click'],
        ['Previous slide', '←, ↑, Page Up, Backspace or P'],
        ['First, last slide', 'Home, End'],
        ['Go to a slide', 'Type its number, then Enter'],
        ['All slides', 'O or Esc'],
        ['Full screen', 'F'],
        ['Black screen', 'B or full stop'],
        ['White screen', 'W or comma'],
      ]
      if (opts.presenter) rows.push(['Presenter view', 'S'])
      rows.push(['This help', '?'])
      var dl = h('dl', { className: 'ap-keys' })
      rows.forEach(function (r) {
        dl.appendChild(h('dt', { text: r[0] }))
        dl.appendChild(h('dd', { text: r[1] }))
      })
      var panel = h('div', { className: 'ap-help-panel' })
      var head = h('div', { className: 'ap-dialog-head' }, [h('h2', { text: 'Keyboard shortcuts' })])
      head.appendChild(button(L, 'close', 'Close (Esc)', function () { closeOverlay() }, 'ap-close'))
      panel.appendChild(head)
      panel.appendChild(dl)
      panel.appendChild(h('p', { className: 'ap-note', text: 'On a touch screen, swipe left or right.' }))
      var el = h('div', { className: 'ap-help', role: 'dialog', 'aria-modal': 'true', 'aria-label': 'Keyboard shortcuts', hidden: true }, [panel])
      L.on(el, 'click', function (e) { if (e.target === el) closeOverlay() })
      return el
    }

    function openOverlay(which) {
      if (overlay) closeOverlay(true)
      if (which === 'overview' && !thumbs) buildOverview()
      var el = which === 'overview' ? overviewEl : helpEl
      lastFocus = document.activeElement
      overlay = which
      el.hidden = false
      root.classList.add('ap-has-overlay')
      hold('overlay')
      if (which === 'overview') {
        markThumb()
        var b = thumbs[index]
        if (b) {
          b.focus()
          if (b.scrollIntoView) b.scrollIntoView({ block: 'center' })
        }
      } else {
        var c = el.querySelector('button')
        if (c) c.focus()
      }
    }

    function closeOverlay(keepFocus) {
      if (!overlay) return
      overviewEl.hidden = true
      helpEl.hidden = true
      overlay = ''
      root.classList.remove('ap-has-overlay')
      release('overlay')
      if (keepFocus !== true) {
        var back = lastFocus && lastFocus !== document.body && root.contains(lastFocus) ? lastFocus : stage
        try { back.focus({ preventScroll: true }) } catch (e) { back.focus() }
      }
    }

    function toggleOverlay(which) {
      if (overlay === which) closeOverlay()
      else openOverlay(which)
    }

    function overviewKey(e) {
      var k = e.key
      if (k === 'Escape' || k === 'Esc' || k === 'o' || k === 'O') {
        e.preventDefault()
        closeOverlay()
        return
      }
      var cur = thumbs.indexOf(document.activeElement)
      if (cur < 0) cur = index
      var cols = 1
      for (var c = 1; c < thumbs.length; c++) {
        if (thumbs[c].offsetTop !== thumbs[0].offsetTop) break
        cols = c + 1
      }
      var to
      if (k === 'ArrowRight' || k === 'Right') to = cur + 1
      else if (k === 'ArrowLeft' || k === 'Left') to = cur - 1
      else if (k === 'ArrowDown' || k === 'Down') to = cur + cols
      else if (k === 'ArrowUp' || k === 'Up') to = cur - cols
      else if (k === 'Home') to = 0
      else if (k === 'End') to = N - 1
      else return
      e.preventDefault()
      var b = thumbs[Math.max(0, Math.min(N - 1, to))]
      b.focus()
      if (b.scrollIntoView) b.scrollIntoView({ block: 'nearest' })
    }

    // ---- presenter window

    function openPresenter() {
      if (!opts.presenter) return
      var open = sync.child()
      if (open) {
        try { open.focus() } catch (e) { /* the browser decides */ }
        return
      }
      var url = location.href.split('#')[0] + '#presenter'
      var win = null
      try { win = window.open(url, NS + '-presenter', 'width=1200,height=760') } catch (e) { win = null }
      if (!win) {
        toast('The browser blocked the presenter window. Allow pop-ups for this page, then press S again.')
        return
      }
      sync.setChild(win)
      // It says hello once it has loaded, and gets the current slide back.
    }

    function toast(text) {
      toastEl.textContent = text
      toastEl.hidden = false
      clearTimeout(toastTimer)
      toastTimer = setTimeout(function () { toastEl.hidden = true }, 5000)
    }

    // ---- input

    function onKey(e) {
      if (e.defaultPrevented || e.ctrlKey || e.metaKey || e.altKey) return
      var t = e.target
      if (isEditable(t)) return
      if (overlay === 'overview') return overviewKey(e)
      if (overlay === 'help') {
        if (e.key === 'Escape' || e.key === 'Esc' || e.key === '?') {
          e.preventDefault()
          closeOverlay()
        }
        return
      }
      var onMedia = closest(t, 'video, audio')
      if (onMedia && e.key !== 'Escape') return

      if (/^[0-9]$/.test(e.key)) {
        numbers.push(e.key)
        e.preventDefault()
        return
      }
      if (numbers.value()) {
        if (e.key === 'Enter') {
          var n = parseInt(numbers.value(), 10)
          numbers.clear()
          e.preventDefault()
          if (n >= 1 && n <= N) go(n - 1)
          return
        }
        if (e.key === 'Backspace') { numbers.back(); e.preventDefault(); return }
        if (e.key === 'Escape' || e.key === 'Esc') { numbers.clear(); e.preventDefault(); return }
      }

      var action = keyAction(e)
      if (!action) return
      // Space and Enter belong to whatever button or link has focus.
      if ((e.key === ' ' || e.key === 'Enter') && closest(t, 'button, a, summary, [role="button"]')) return
      e.preventDefault()
      if (blank && action !== 'black' && action !== 'white') {
        setBlank('')
        return
      }
      switch (action) {
        case 'next': next(); break
        case 'prev': prev(); break
        case 'first': go(nav.first(), { forward: false }); break
        case 'last': go(nav.last(), { forward: true }); break
        case 'full': toggleFullscreen(); break
        case 'black': setBlank(blank === 'black' ? '' : 'black'); break
        case 'white': setBlank(blank === 'white' ? '' : 'white'); break
        case 'overview': case 'escape': toggleOverlay('overview'); break
        case 'presenter': openPresenter(); break
        case 'help': toggleOverlay('help'); break
      }
    }

    function onStageClick(e) {
      if (e.button !== 0 || now() - swipedAt < 500) return
      if (closest(e.target, 'a, button, video, audio, input, select, textarea, label')) return
      next()
    }

    function onLinkClick(e) {
      var a = closest(e.target, 'a[data-slide]')
      if (!a || e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey) return
      e.preventDefault()
      go(parseInt(a.getAttribute('data-slide'), 10))
    }

    function onTouchStart(e) {
      if (e.touches.length !== 1 || closest(e.target, 'a, button, video, audio')) {
        touch = null
        return
      }
      touch = { x: e.touches[0].clientX, y: e.touches[0].clientY }
    }

    function onTouchEnd(e) {
      if (!touch || !e.changedTouches.length) return
      var dx = e.changedTouches[0].clientX - touch.x
      var dy = e.changedTouches[0].clientY - touch.y
      touch = null
      if (Math.abs(dx) < 50 || Math.abs(dx) < Math.abs(dy) * 1.5) return
      swipedAt = now()
      if (dx < 0) next()
      else prev()
    }

    // Controls and cursor fade when the mouse rests — only where there is a
    // mouse; on a touch screen there'd be no way to get them back.
    function wake() {
      root.classList.remove('ap-idle')
      clearTimeout(idleTimer)
      idleTimer = setTimeout(sleep, IDLE_MS)
    }

    function sleep() {
      var focusInBar = bar && bar.contains(document.activeElement)
      if (overlay || barHover || focusInBar) {
        idleTimer = setTimeout(sleep, IDLE_MS)
        return
      }
      if (mq('(hover: hover)')) root.classList.add('ap-idle')
    }

    function fit() {
      var W = stage.clientWidth
      var H = stage.clientHeight
      if (!W || !H) return
      var w = Math.min(W, H * data.aspect)
      frame.style.width = Math.floor(w) + 'px'
      frame.style.height = Math.floor(w / data.aspect) + 'px'
    }

    function onFullscreen() {
      var on = !!fullscreenElement()
      setButton(btn.full, on ? 'unfull' : 'full', on ? 'Leave full screen (F)' : 'Full screen (F)')
      fit()
    }

    L.on(document, 'keydown', onKey)
    L.on(stage, 'click', onStageClick)
    L.on(root, 'click', onLinkClick)
    L.on(blankEl, 'click', function () { setBlank('') })
    L.on(stage, 'touchstart', onTouchStart, { passive: true })
    L.on(stage, 'touchend', onTouchEnd, { passive: true })
    L.on(document, 'mousemove', wake)
    L.on(document, 'focusin', function (e) { if (bar && bar.contains(e.target)) wake() })
    L.on(window, 'hashchange', onHash)
    L.on(window, 'popstate', onHash)
    L.on(window, 'resize', fit)
    L.on(document, 'fullscreenchange', onFullscreen)
    L.on(document, 'webkitfullscreenchange', onFullscreen)
    var ro = null
    if (typeof ResizeObserver === 'function') {
      ro = new ResizeObserver(fit)
      ro.observe(stage)
    }

    fit()
    wake()
    var first = hashIndex()
    go(first >= 0 ? first : nav.first(), { forward: true })

    var ctrl = {
      mode: 'audience',
      index: function () { return index },
      count: N,
      go: function (i) { go(i) },
      next: next,
      prev: prev,
      blank: function () { return blank },
      overlay: function () { return overlay },
      destroy: function () {
        L.off()
        if (ro) ro.disconnect()
        clearTimeout(auto.timer)
        clearTimeout(idleTimer)
        clearTimeout(toastTimer)
        numbers.stop()
        settle()
        pauseMedia(view)
        sync.close()
        Object.keys(blobs).forEach(function (k) {
          try { URL.revokeObjectURL(blobs[k]) } catch (e) { /* gone */ }
        })
        if (root.parentNode) root.parentNode.removeChild(root)
        document.documentElement.classList.remove('ap-on')
        if (api.current === ctrl) api.current = null
      },
    }
    return ctrl
  }

  // ---------------------------------------------------------------- presenter

  function createPresenter(data, host) {
    var L = new Listeners()
    var slides = data.slides
    var opts = data.options
    var nav = makeNav(slides, opts)
    var N = nav.N
    var index = -1
    var blank = ''
    var clockTimer = 0
    var timer = { start: now(), pausedAt: 0 }
    var notesPx = Math.max(12, Math.min(48, parseInt(store('notes-px'), 10) || 20))

    document.documentElement.classList.add('ap-on')
    document.title = 'Presenter view' + (data.title ? ' — ' + data.title : '')
    var root = h('div', { className: 'ap ap-presenter', 'data-theme': opts.theme })
    root.style.setProperty('--ap-slide-bg', data.background)

    var elapsed = h('span', { className: 'pv-elapsed', role: 'timer', 'aria-label': 'Time elapsed', text: '0:00' })
    var clock = h('span', { className: 'pv-clock', 'aria-label': 'Time of day' })
    var pauseBtn = button(L, 'pause', 'Pause the timer', toggleTimer)
    var resetBtn = button(L, 'reset', 'Reset the timer', resetTimer)
    var head = h('header', { className: 'pv-head' }, [
      h('h1', { text: data.title || 'Presenter view' }),
      h('div', { className: 'pv-timer' }, [elapsed, pauseBtn, resetBtn]),
      clock,
    ])

    var curImg = h('img', { alt: '' })
    var curLabel = h('h2', { className: 'pv-label' })
    var cur = h('section', { className: 'pv-current', 'aria-label': 'Current slide' }, [curLabel, h('div', { className: 'pv-box' }, [curImg])])

    var nextImg = h('img', { alt: '' })
    var nextEnd = h('p', { className: 'pv-end', text: 'End of the presentation' })
    var nextLabel = h('h2', { className: 'pv-label', text: 'Next' })
    var nxt = h('section', { className: 'pv-next', 'aria-label': 'Next slide' }, [nextLabel, h('div', { className: 'pv-box' }, [nextImg, nextEnd])])

    var notesBody = h('div', { className: 'pv-notes-body', tabindex: '0', 'aria-label': 'Speaker notes' })
    var notes = h('section', { className: 'pv-notes', 'aria-label': 'Notes' }, [
      h('div', { className: 'pv-notes-head' }, [
        h('h2', { className: 'pv-label', text: 'Notes' }),
        button(L, 'smaller', 'Smaller notes text', function () { setNotesPx(notesPx - 2) }),
        button(L, 'bigger', 'Larger notes text', function () { setNotesPx(notesPx + 2) }),
      ]),
      notesBody,
    ])

    var counter = h('span', { className: 'ap-counter' })
    var status = h('span', { className: 'pv-status', role: 'status' })
    var blankBtn = button(L, 'blank', 'Black screen (B)', function () { setBlank(blank === 'black' ? '' : 'black') })
    var foot = h('footer', { className: 'pv-foot' }, [
      button(L, 'prev', 'Previous slide', prev),
      counter,
      button(L, 'next', 'Next slide', next),
      blankBtn,
      status,
    ])
    var live = h('div', { className: 'ap-sr', 'aria-live': 'polite', 'aria-atomic': 'true' })

    root.appendChild(head)
    root.appendChild(cur)
    root.appendChild(nxt)
    root.appendChild(notes)
    root.appendChild(foot)
    root.appendChild(live)
    host.appendChild(root)

    var numbers = makeNumberEntry(function (d) {
      setStatus(d ? 'Go to slide ' + d + ' ↵' : statusText())
    })

    var sync = makeSync(deckKey(data), function (st) {
      if (typeof st.index === 'number') go(st.index, true)
      if (typeof st.blank === 'string') setBlank(st.blank, true)
    }, function () {
      sync.post('state', { index: index, blank: blank })
    }, L)

    function go(i, remote) {
      if (typeof i !== 'number' || isNaN(i)) return
      i = Math.max(0, Math.min(N - 1, Math.floor(i)))
      if (i !== index) {
        index = i
        render()
      }
      if (!remote) sync.post('state', { index: index, blank: blank })
    }

    function next() {
      if (blank) return setBlank('')
      var j = nav.step(index, 1)
      if (j >= 0) go(j)
    }

    function prev() {
      if (blank) return setBlank('')
      var j = nav.step(index, -1)
      if (j >= 0) go(j)
    }

    function setBlank(mode, remote) {
      mode = mode === 'black' || mode === 'white' ? mode : ''
      if (mode === blank) return
      blank = mode
      root.classList.toggle('pv-blanked', !!mode)
      setStatus(statusText())
      if (!remote) sync.post('state', { index: index, blank: blank })
    }

    function render() {
      var s = slides[index]
      curImg.src = resolveSrc(s.src)
      curImg.alt = s.alt || s.title || 'Slide ' + (index + 1)
      curLabel.textContent = 'Slide ' + (index + 1) + ' of ' + N + (s.hidden ? ' (hidden)' : '') + (s.title ? ' — ' + s.title : '')
      counter.textContent = index + 1 + ' / ' + N
      var j = nav.step(index, 1)
      if (j >= 0) {
        nextImg.hidden = false
        nextEnd.hidden = true
        nextImg.src = resolveSrc(slides[j].src)
        nextImg.alt = slides[j].alt || slides[j].title || 'Slide ' + (j + 1)
        nextLabel.textContent = 'Next: slide ' + (j + 1)
      } else {
        nextImg.hidden = true
        nextImg.removeAttribute('src')
        nextEnd.hidden = false
        nextLabel.textContent = 'Next'
      }
      // Notes are HTML the exporter built, with everything from the deck escaped.
      if (s.notes) notesBody.innerHTML = s.notes
      else notesBody.innerHTML = '<p class="pv-none">No notes for this slide.</p>'
      notesBody.scrollTop = 0
      live.textContent = slideLabel(s, index, N)
    }

    function setNotesPx(px) {
      notesPx = Math.max(12, Math.min(48, px))
      notesBody.style.fontSize = notesPx + 'px'
      store('notes-px', notesPx)
    }

    // Only touched when it changes: it is a live region, and the clock ticks.
    function setStatus(text) {
      if (status.textContent !== text) status.textContent = text
    }

    function statusText() {
      var parts = []
      if (blank) parts.push('The audience sees a ' + blank + ' screen')
      parts.push(sync.linked() ? 'Linked to the audience window' : 'Not linked to an audience window')
      return parts.join(' · ')
    }

    function elapsedMs() {
      return (timer.pausedAt || now()) - timer.start
    }

    function toggleTimer() {
      if (timer.pausedAt) {
        timer.start += now() - timer.pausedAt
        timer.pausedAt = 0
      } else {
        timer.pausedAt = now()
      }
      setButton(pauseBtn, timer.pausedAt ? 'play' : 'pause', timer.pausedAt ? 'Restart the timer' : 'Pause the timer')
      root.classList.toggle('pv-timer-paused', !!timer.pausedAt)
      tick()
    }

    function resetTimer() {
      timer.start = now()
      if (timer.pausedAt) timer.pausedAt = timer.start
      tick()
    }

    function tick() {
      var s = Math.floor(elapsedMs() / 1000)
      var hh = Math.floor(s / 3600)
      var mm = Math.floor((s % 3600) / 60)
      var ss = s % 60
      elapsed.textContent = (hh ? hh + ':' + (mm < 10 ? '0' : '') : '') + mm + ':' + (ss < 10 ? '0' : '') + ss
      try {
        clock.textContent = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
      } catch (e) {
        clock.textContent = new Date().toTimeString().slice(0, 5)
      }
      if (!numbers.value()) setStatus(statusText())
    }

    function onKey(e) {
      if (e.defaultPrevented || e.ctrlKey || e.metaKey || e.altKey) return
      var t = e.target
      if (isEditable(t)) return
      if (/^[0-9]$/.test(e.key)) {
        numbers.push(e.key)
        e.preventDefault()
        return
      }
      if (numbers.value()) {
        if (e.key === 'Enter') {
          var n = parseInt(numbers.value(), 10)
          numbers.clear()
          e.preventDefault()
          if (n >= 1 && n <= N) go(n - 1)
          return
        }
        if (e.key === 'Backspace') { numbers.back(); e.preventDefault(); return }
        if (e.key === 'Escape' || e.key === 'Esc') { numbers.clear(); e.preventDefault(); return }
      }
      // In the notes, the arrows, Space and Home/End scroll the notes. A
      // clicker's Page Up/Down still moves the slides.
      if (t === notesBody && /^(ArrowUp|ArrowDown|Up|Down| |Spacebar|Home|End)$/.test(e.key)) return
      var action = keyAction(e)
      if (!action) return
      if ((e.key === ' ' || e.key === 'Enter') && closest(t, 'button, a')) return
      e.preventDefault()
      switch (action) {
        case 'next': next(); break
        case 'prev': prev(); break
        case 'first': go(nav.first()); break
        case 'last': go(nav.last()); break
        case 'full': toggleFullscreen(); break
        case 'black': setBlank(blank === 'black' ? '' : 'black'); break
        case 'white': setBlank(blank === 'white' ? '' : 'white'); break
      }
    }

    L.on(document, 'keydown', onKey)
    setNotesPx(notesPx)
    index = Math.max(0, nav.first())
    render()
    tick()
    clockTimer = setInterval(tick, 1000)
    sync.post('hello')

    var ctrl = {
      mode: 'presenter',
      index: function () { return index },
      count: N,
      go: function (i) { go(i) },
      next: next,
      prev: prev,
      blank: function () { return blank },
      overlay: function () { return '' },
      destroy: function () {
        L.off()
        clearInterval(clockTimer)
        numbers.stop()
        sync.close()
        if (root.parentNode) root.parentNode.removeChild(root)
        document.documentElement.classList.remove('ap-on')
        if (api.current === ctrl) api.current = null
      },
    }
    return ctrl
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { start() })
  } else {
    start()
  }
})()
