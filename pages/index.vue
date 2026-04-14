<template>
  <div id="app">

    <!-- HEADER -->
    <header :class="{scrolled: scrolled}">
      <div class="hdr">
        <a class="brand" @click="goto('hero')">
          <span class="brand-name">Nucleo<span class="brand-blue">us</span></span>
        </a>
        <nav class="nav-links">
          <div class="nav-item" @mouseenter="activeDropdown='product'" @mouseleave="activeDropdown=null">
            <a class="nav-link">Product <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg></a>
            <div class="dropdown" v-if="activeDropdown==='product'">
              <div class="dd-section">
                <div class="dd-head">Platform</div>
                <a class="dd-item" @click="goto('features')"><div class="dd-ico"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg></div><div><div class="dd-title">Pipeline Automation</div><div class="dd-sub">Automate lead stages and transitions</div></div></a>
                <a class="dd-item" @click="goto('features')"><div class="dd-ico"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg></div><div><div class="dd-title">AI Lead Scoring</div><div class="dd-sub">Score and qualify leads automatically</div></div></a>
                <a class="dd-item" @click="goto('features')"><div class="dd-ico"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 17V9m4 8v-5m4 5v-3"/></svg></div><div><div class="dd-title">Analytics & Reporting</div><div class="dd-sub">Real-time funnel and revenue insights</div></div></a>
              </div>
              <div class="dd-section">
                <div class="dd-head">Workspace</div>
                <a class="dd-item" @click="goto('features')"><div class="dd-ico"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/></svg></div><div><div class="dd-title">Team Management</div><div class="dd-sub">Assign leads and track performance</div></div></a>
                <a class="dd-item" @click="goto('features')"><div class="dd-ico"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></div><div><div class="dd-title">Quotation Builder</div><div class="dd-sub">Generate quotes in minutes</div></div></a>
                <a class="dd-item" @click="openViewProduct"><div class="dd-ico" style="background:rgba(37,99,235,0.12)"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8m-4-4v4"/></svg></div><div><div class="dd-title" style="color:#2563eb">View Live Product</div><div class="dd-sub">Access employer or admin panel</div></div></a>
              </div>
            </div>
          </div>
          <div class="nav-item" @mouseenter="activeDropdown='solutions'" @mouseleave="activeDropdown=null">
            <a class="nav-link">Solutions <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg></a>
            <div class="dropdown" v-if="activeDropdown==='solutions'">
              <div class="dd-section">
                <a class="dd-item" v-for="s in solutions" :key="s.title"><div class="dd-ico"><span v-html="s.svg"></span></div><div><div class="dd-title">{{ s.title }}</div><div class="dd-sub">{{ s.sub }}</div></div></a>
              </div>
            </div>
          </div>
          <a class="nav-link" @click="goto('about')">About</a>
          <a class="nav-link" @click="goto('how-it-works')">How It Works</a>
          <a class="nav-link" @click="goto('contact')">Contact</a>
        </nav>
        <div class="hdr-right">
          <button class="view-product-btn" @click="openViewProduct">View Product</button>
          <button class="ghost-btn" @click="showLogin=true">Sign In</button>
          <button class="book-btn" @click="showSignup=true">Get Started</button>
          <button class="mob-menu" @click="mobOpen=!mobOpen">
            <span :class="{o1:mobOpen}"></span>
            <span :class="{o2:mobOpen}"></span>
            <span :class="{o3:mobOpen}"></span>
          </button>
        </div>
      </div>
      <Transition name="mob-drop">
        <div v-if="mobOpen" class="mob-nav">
          <a @click="goto('about');mobOpen=false">About</a>
          <a @click="goto('features');mobOpen=false">Features</a>
          <a @click="goto('how-it-works');mobOpen=false">How It Works</a>
          <a @click="goto('contact');mobOpen=false">Contact</a>
          <a @click="openViewProduct();mobOpen=false">View Product</a>
          <div class="mob-auth-row">
            <button class="ghost-btn w50" @click="showLogin=true;mobOpen=false">Sign In</button>
            <button class="book-btn w50" @click="showSignup=true;mobOpen=false">Get Started</button>
          </div>
        </div>
      </Transition>
    </header>

    <!-- HERO -->
    <section id="hero" class="hero-section">
      <div class="hero-bg">
        <div class="hero-grid-bg"></div>
        <div class="hero-overlay"></div>
      </div>
      <div class="container hero-content">
        <div class="hero-badge">
          <span class="badge-dot"></span>
          AI-Powered Lead Operations — Built for Scale
        </div>
        <h1 class="hero-title">
          One Platform.<br>
          Every Lead. Every Team.<br>
          <span class="blue-text">Every Insight.</span>
        </h1>
        <p class="hero-sub">Nucleous replaces disconnected tools with a single intelligent system — capturing leads, orchestrating pipelines, automating follow-ups, and delivering workforce intelligence in real time.</p>
        <div class="hero-btns">
          <button class="book-btn lg" @click="showSignup=true">Start Free Trial</button>
          <button class="ghost-btn lg" @click="openViewProduct">View Live Product</button>
        </div>
        <div class="hero-trust-row">
          <div class="ht-item" v-for="t in trustStats" :key="t.label">
            <span class="ht-num">{{ t.num }}</span>
            <span class="ht-lbl">{{ t.label }}</span>
          </div>
        </div>
      </div>

      <div class="hero-dash container">
        <div class="dash-window">
          <div class="dash-topbar">
            <span class="dtb-dot r"></span><span class="dtb-dot y"></span><span class="dtb-dot g"></span>
            <span class="dtb-title">Nucleous — Operations Dashboard</span>
            <span class="dtb-live">Live</span>
          </div>
          <div class="dash-body">
            <div class="dash-side">
              <div class="dsi active"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 17V9m4 8v-5m4 5v-3"/></svg><span>Overview</span></div>
              <div class="dsi" v-for="s in ['Pipeline','Leads','Team','Quotations','Analytics','Reports','Settings']" :key="s">{{ s }}</div>
            </div>
            <div class="dash-main">
              <div class="dash-kpi-row">
                <div class="kpi" v-for="k in kpis" :key="k.label">
                  <div class="kpi-val" :style="{color:k.color}">{{ k.val }}</div>
                  <div class="kpi-label">{{ k.label }}</div>
                  <div class="kpi-chg" :class="k.up ? 'up':'dn'">{{ k.chg }}</div>
                </div>
              </div>
              <div class="dash-charts-row">
                <div class="dash-chart-box">
                  <div class="dcb-label">Lead Funnel — Last 30 Days</div>
                  <svg viewBox="0 0 220 80" preserveAspectRatio="none" style="width:100%;height:70px">
                    <defs><linearGradient id="cgrad1" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#2563eb" stop-opacity=".18"/><stop offset="100%" stop-color="#2563eb" stop-opacity="0"/></linearGradient></defs>
                    <path d="M0 65 C30 58 50 44 75 32 C100 20 115 38 140 18 C165 4 185 16 220 6" fill="none" stroke="#2563eb" stroke-width="2" stroke-linecap="round"/>
                    <path d="M0 65 C30 58 50 44 75 32 C100 20 115 38 140 18 C165 4 185 16 220 6 L220 80 L0 80Z" fill="url(#cgrad1)"/>
                    <circle cx="140" cy="18" r="3" fill="#2563eb"/><circle cx="220" cy="6" r="3" fill="#2563eb"/>
                  </svg>
                </div>
                <div class="dash-chart-box">
                  <div class="dcb-label">Closure Rate by Stage</div>
                  <div class="mini-bars">
                    <div class="mb-item" v-for="b in miniBars" :key="b.label">
                      <div class="mb-bar-wrap"><div class="mb-bar" :style="{width:b.pct+'%',background:b.color}"></div></div>
                      <div class="mb-label">{{ b.label }}</div>
                      <div class="mb-val">{{ b.pct }}%</div>
                    </div>
                  </div>
                </div>
                <div class="dash-chart-box">
                  <div class="dcb-label">Team Activity Today</div>
                  <div class="team-activity">
                    <div class="ta-row" v-for="t in teamActivity" :key="t.name">
                      <div class="ta-avatar">{{ t.initial }}</div>
                      <div class="ta-info"><div class="ta-name">{{ t.name }}</div><div class="ta-stat">{{ t.stat }}</div></div>
                      <div class="ta-badge" :class="t.status">{{ t.statusLabel }}</div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- METRICS STRIP -->
    <div class="metrics-strip">
      <div class="container ms-inner">
        <div class="ms-item" v-for="m in metricsStrip" :key="m.label">
          <div class="ms-num">{{ m.num }}</div>
          <div class="ms-label">{{ m.label }}</div>
        </div>
      </div>
    </div>

    <!-- LOGOS STRIP -->
    <div class="logos-strip">
      <div class="container ls-inner">
        <span class="ls-label">Integrates with</span>
        <div class="ls-logos">
          <span v-for="l in integrations" :key="l" class="ls-logo">{{ l }}</span>
        </div>
      </div>
    </div>

    <!-- ABOUT -->
    <section id="about" class="section">
      <div class="container two-col">
        <div>
          <div class="eyebrow">What is Nucleous</div>
          <h2 class="section-title">The Operational Core of Every Lead-Driven Organization</h2>
          <p class="body-text">Most businesses manage leads in spreadsheets, follow up over WhatsApp, and track pipelines in separate tools. Nucleous brings every piece together — from first touchpoint to final closure — into one structured, AI-guided environment.</p>
          <div class="about-points">
            <div class="apt" v-for="p in aboutPoints" :key="p.title">
              <div class="apt-icon" v-html="p.svg"></div>
              <div>
                <div class="apt-title">{{ p.title }}</div>
                <div class="apt-desc">{{ p.desc }}</div>
              </div>
            </div>
          </div>
          <button class="book-btn" @click="showSignup=true">Start Free Trial</button>
        </div>
        <div class="about-visual">
          <div class="av-stats-grid">
            <div class="avs-card" v-for="s in aboutStats" :key="s.label" :style="{borderTopColor:s.color}">
              <div class="avs-num" :style="{color:s.color}">{{ s.num }}</div>
              <div class="avs-label">{{ s.label }}</div>
              <div class="avs-desc">{{ s.desc }}</div>
            </div>
          </div>
          <div class="av-card dark-card">
            <div class="avc-label">ZIRA — AI Assistant for Nucleous</div>
            <p class="avc-desc">ZIRA is Nucleous's built-in AI, trained exclusively on your pipeline data. It guides follow-ups, recommends actions, forecasts outcomes, and answers questions about your leads and team — in real time.</p>
            <button class="avc-btn" @click="showSignup=true">Talk to ZIRA</button>
          </div>
        </div>
      </div>
    </section>

    <!-- FEATURES -->
    <section id="features" class="section gray-bg">
      <div class="container">
        <div class="section-center">
          <div class="eyebrow">Core Capabilities</div>
          <h2 class="section-title">Built for the Full Revenue Cycle</h2>
          <p class="body-text centered-text">From the first lead capture to the final invoice — every function purpose-built and connected.</p>
        </div>
        <div class="feat-grid">
          <div class="feat-card" v-for="f in features" :key="f.title">
            <div class="fc-icon" v-html="f.svg"></div>
            <div class="fc-title">{{ f.title }}</div>
            <div class="fc-desc">{{ f.desc }}</div>
          </div>
        </div>
      </div>
    </section>

    <!-- ANALYTICS SECTION -->
    <section class="section analytics-section">
      <div class="container">
        <div class="section-center">
          <div class="eyebrow">Intelligence Layer</div>
          <h2 class="section-title">Data That Drives Decisions</h2>
          <p class="body-text centered-text">Nucleous surfaces trends, bottlenecks, and revenue opportunities across every funnel stage and team.</p>
        </div>
        <div class="analytics-grid">
          <div class="an-card large">
            <div class="an-card-hdr"><div class="an-card-title">Revenue Pipeline — 12 Months</div><div class="an-card-sub">Total managed value</div></div>
            <div class="an-big-num">$2.4M <span class="an-chg up">+28%</span></div>
            <svg viewBox="0 0 480 100" preserveAspectRatio="none" style="width:100%;height:90px;margin-top:8px">
              <defs><linearGradient id="rgrad" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#2563eb" stop-opacity=".2"/><stop offset="100%" stop-color="#2563eb" stop-opacity="0"/></linearGradient></defs>
              <path d="M0 90 C40 80 60 60 90 45 C120 30 140 55 180 35 C220 15 240 28 280 15 C320 4 350 18 390 10 C420 4 450 8 480 3" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round"/>
              <path d="M0 90 C40 80 60 60 90 45 C120 30 140 55 180 35 C220 15 240 28 280 15 C320 4 350 18 390 10 C420 4 450 8 480 3 L480 100 L0 100Z" fill="url(#rgrad)"/>
            </svg>
            <div class="an-legend-row">
              <div class="an-leg" v-for="l in revLegend" :key="l.label"><span class="al-dot" :style="{background:l.color}"></span>{{ l.label }}: <strong>{{ l.val }}</strong></div>
            </div>
          </div>
          <div class="an-card">
            <div class="an-card-hdr"><div class="an-card-title">Lead Source Breakdown</div></div>
            <div class="donut-wrap">
              <svg viewBox="0 0 120 120" style="width:110px;height:110px">
                <circle cx="60" cy="60" r="45" fill="none" stroke="#e8eaf0" stroke-width="18"/>
                <circle cx="60" cy="60" r="45" fill="none" stroke="#2563eb" stroke-width="18" stroke-dasharray="113 170" stroke-dashoffset="0" stroke-linecap="round"/>
                <circle cx="60" cy="60" r="45" fill="none" stroke="#059669" stroke-width="18" stroke-dasharray="68 215" stroke-dashoffset="-113" stroke-linecap="round"/>
                <circle cx="60" cy="60" r="45" fill="none" stroke="#7c3aed" stroke-width="18" stroke-dasharray="51 232" stroke-dashoffset="-181" stroke-linecap="round"/>
                <text x="60" y="64" text-anchor="middle" font-size="14" font-weight="700" fill="#0a0b14" font-family="'DM Sans',sans-serif">342</text>
              </svg>
              <div class="donut-legend">
                <div class="dl-item" v-for="d in donutData" :key="d.label"><span class="al-dot" :style="{background:d.color}"></span><span>{{ d.label }}</span><strong>{{ d.pct }}%</strong></div>
              </div>
            </div>
          </div>
          <div class="an-card">
            <div class="an-card-hdr"><div class="an-card-title">Avg Response Time</div></div>
            <div class="an-big-num blue">4.2m <span class="an-chg up">-18%</span></div>
            <div class="resp-bars">
              <div class="rb-item" v-for="r in respBars" :key="r.day">
                <div class="rb-bar-wrap"><div class="rb-bar" :style="{height:r.h+'px',background:r.active ? '#2563eb' : '#dbeafe'}"></div></div>
                <div class="rb-label">{{ r.day }}</div>
              </div>
            </div>
          </div>
          <div class="an-card">
            <div class="an-card-hdr"><div class="an-card-title">AI Follow-up Performance</div></div>
            <div class="ai-perf-rows">
              <div class="ap-row" v-for="a in aiPerf" :key="a.label">
                <div class="ap-label">{{ a.label }}</div>
                <div class="ap-bar-wrap"><div class="ap-bar" :style="{width:a.pct+'%',background:a.color}"></div></div>
                <div class="ap-val">{{ a.pct }}%</div>
              </div>
            </div>
          </div>
          <div class="an-card">
            <div class="an-card-hdr"><div class="an-card-title">Funnel Velocity</div></div>
            <div class="funnel-stages">
              <div class="fs-row" v-for="fs in funnelStages" :key="fs.name">
                <div class="fs-name">{{ fs.name }}</div>
                <div class="fs-bar-wrap"><div class="fs-bar" :style="{width:fs.pct+'%'}"></div></div>
                <div class="fs-count">{{ fs.count }}</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- HOW IT WORKS -->
    <section id="how-it-works" class="section gray-bg">
      <div class="container">
        <div class="section-center">
          <div class="eyebrow">How It Works</div>
          <h2 class="section-title">From First Lead to Final Closure</h2>
          <p class="body-text centered-text">Nucleous structures every stage of your revenue operation — no lead left unattended, no action without context.</p>
        </div>
        <div class="steps-row">
          <div class="step-card" v-for="(s,i) in steps" :key="s.title" :class="{active: activeStep===i}" @click="activeStep=i">
            <div class="sc-num">{{ String(i+1).padStart(2,'0') }}</div>
            <div class="sc-title">{{ s.title }}</div>
            <div class="sc-desc">{{ s.desc }}</div>
            <ul class="sc-list" v-if="activeStep===i">
              <li v-for="pt in s.points" :key="pt">{{ pt }}</li>
            </ul>
          </div>
        </div>
      </div>
    </section>

    <!-- ZIRA SECTION -->
    <section class="section zira-section">
      <div class="container two-col">
        <div class="zira-visual">
          <div class="zira-chat-mock">
            <div class="zcm-hdr">
              <div class="zcm-avatar">Z</div>
              <div><div class="zcm-name">ZIRA</div><div class="zcm-status">AI Assistant — Nucleous</div></div>
            </div>
            <div class="zcm-messages">
              <div class="zcm-msg bot">Good morning. You have 12 uncontacted leads from the last 24 hours. Shall I schedule follow-ups?</div>
              <div class="zcm-msg user">Yes, prioritize by lead score.</div>
              <div class="zcm-msg bot">Done. 8 high-priority leads assigned to your team with personalized WhatsApp messages queued. Top lead: Arjun Sharma — Intent score 94.</div>
              <div class="zcm-msg user">What's our closure rate this week?</div>
              <div class="zcm-msg bot">This week: 34 deals closed, 89% closure rate — up 12% from last week. Revenue attributed: 318,000. Your top performer is Priya Mehta with 11 closures.</div>
            </div>
            <div class="zcm-input">Ask ZIRA anything about your pipeline...</div>
          </div>
        </div>
        <div>
          <div class="eyebrow">Meet ZIRA</div>
          <h2 class="section-title">Your AI Copilot, Trained on Nucleous</h2>
          <p class="body-text">ZIRA is not a generic chatbot. It is trained specifically on Nucleous workflows, your pipeline data, team activity, and lead history. It reasons about your business and acts — not just responds.</p>
          <div class="zira-caps">
            <div class="zc-item" v-for="z in ziraCaps" :key="z.title">
              <div class="zc-dot"></div>
              <div><div class="zc-title">{{ z.title }}</div><div class="zc-desc">{{ z.desc }}</div></div>
            </div>
          </div>
          <button class="book-btn" @click="showSignup=true">Experience ZIRA</button>
        </div>
      </div>
    </section>

    <!-- CONTACT -->
    <section id="contact" class="section gray-bg">
      <div class="container two-col">
        <div>
          <div class="eyebrow">Get In Touch</div>
          <h2 class="section-title">Ready to Unify Your Sales Operations?</h2>
          <p class="body-text">Book a 30-minute walkthrough and see how Nucleous transforms lead operations for your team.</p>
          <div class="contact-info">
            <div class="ci-row" v-for="c in contactInfo" :key="c.label">
              <div class="ci-ico" v-html="c.svg"></div>
              <div><div class="ci-lbl">{{ c.label }}</div><div class="ci-val">{{ c.val }}</div></div>
            </div>
          </div>
        </div>
        <div class="contact-form-wrap">
          <div v-if="contactDone" class="form-success">
            <div class="success-icon">&#10003;</div>
            <h4>Message Received</h4>
            <p>We will get back to you within 24 hours. Check your email for confirmation.</p>
          </div>
          <form v-else @submit.prevent="submitContact" class="contact-form">
            <div class="cf-row-2">
              <div class="cf-g">
                <label>Full Name</label>
                <input v-model="cf.name" placeholder="Your name" required @blur="trackHalfFill('contact')"/>
              </div>
              <div class="cf-g">
                <label>Work Email</label>
                <input v-model="cf.email" type="email" placeholder="you@company.com" required @blur="trackHalfFill('contact')"/>
              </div>
            </div>
            <div class="cf-g">
              <label>Company</label>
              <input v-model="cf.company" placeholder="Company name" @blur="trackHalfFill('contact')"/>
            </div>
            <div class="cf-g">
              <label>Team Size</label>
              <select v-model="cf.size" @blur="trackHalfFill('contact')">
                <option value="">Select size</option>
                <option>1-10</option><option>11-50</option><option>51-200</option><option>200+</option>
              </select>
            </div>
            <div class="cf-g">
              <label>Message</label>
              <textarea v-model="cf.message" rows="3" placeholder="Tell us about your sales operations challenge..." @blur="trackHalfFill('contact')"></textarea>
            </div>
            <button type="submit" class="book-btn w100" :disabled="contactSending">{{ contactSending ? 'Sending...' : 'Book a Demo' }}</button>
          </form>
        </div>
      </div>
    </section>

    <!-- CTA BANNER -->
    <section class="cta-section">
      <div class="container cta-inner">
        <div class="cta-left">
          <h2 class="cta-title">Start Managing Leads Intelligently</h2>
          <p class="cta-sub">Join teams that have moved from fragmented tools to a single source of revenue truth.</p>
        </div>
        <div class="cta-right">
          <button class="cta-btn-w" @click="showSignup=true">Create Free Account</button>
          <button class="cta-btn-o" @click="showLogin=true">Sign In to Dashboard</button>
          <p class="cta-note">No credit card required · 14-day free trial</p>
        </div>
      </div>
    </section>

    <!-- FOOTER -->
    <footer>
      <div class="container foot-grid">
        <div class="foot-brand-col">
          <div class="brand"><span class="brand-name">nucle<span class="brand-blue">us</span></span></div>
          <p class="foot-desc">AI-driven lead operations and workforce intelligence platform for modern revenue teams.</p>
        </div>
        <div class="foot-links-grid">
          <div class="fl-col" v-for="col in footLinks" :key="col.head">
            <div class="fl-head">{{ col.head }}</div>
            <a v-for="lnk in col.links" :key="lnk" @click="handleFootLink(lnk)">{{ lnk }}</a>
          </div>
        </div>
      </div>
      <div class="foot-bottom">
        <div class="container fbt-inner">
          <span>2025 Nucleous. All rights reserved.</span>
          <span>Lead Operations · AI Automation · Workforce Intelligence</span>
        </div>
      </div>
    </footer>

    <!-- LOGIN MODAL -->
    <Transition name="modal">
      <div v-if="showLogin" class="modal-overlay" @click.self="closeLogin">
        <div class="modal-box">
          <div class="modal-hdr">
            <div class="modal-brand"><span class="brand-name sm">nucle<span class="brand-blue">us</span></span></div>
            <button class="modal-close" @click="closeLogin">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </button>
          </div>

          <!-- STEP 0: Credentials -->
          <template v-if="loginStep === 0">
            <h3 class="modal-title">Welcome back</h3>
            <p class="modal-sub">Sign in to your Nucleous account</p>
            <form @submit.prevent="submitLogin" class="auth-form">
              <div v-if="loginError" class="login-error">{{ loginError }}</div>
              <div class="af-g">
                <label>Email Address</label>
                <input v-model="login.email" type="email" placeholder="you@company.com" required/>
              </div>
              <div class="af-g">
                <div class="af-lbl-row">
                  <label>Password</label>
                  <a class="forgot-link">Forgot password?</a>
                </div>
                <input v-model="login.password" type="password" placeholder="Enter your password" required/>
              </div>
              <button type="submit" class="book-btn w100 lg" :disabled="loginLoading">
                {{ loginLoading ? 'Sending OTP...' : 'Continue — Send OTP' }}
              </button>
              <p class="auth-switch">Don't have an account? <a @click="showLogin=false;showSignup=true">Create one free</a></p>
            </form>
          </template>

          <!-- STEP 1: OTP -->
          <template v-else-if="loginStep === 1">
            <div class="otp-header-icon">
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="1.8"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
            </div>
            <h3 class="modal-title">Check Your Email</h3>
            <p class="modal-sub">We sent a 6-digit code to <strong>{{ login.email }}</strong></p>
            <form @submit.prevent="verifyOTP" class="auth-form">
              <div v-if="otpError" class="login-error">{{ otpError }}</div>
              <div class="af-g otp-g">
                <label>Enter Verification Code</label>
                <input v-model="otpValue" type="text" inputmode="numeric" pattern="[0-9]*" placeholder="000000" maxlength="6" class="otp-input" required autocomplete="one-time-code"/>
                <div class="otp-hint">Enter the 6-digit code sent to your email</div>
              </div>
              <button type="submit" class="book-btn w100 lg" :disabled="otpLoading || otpValue.length < 6">
                {{ otpLoading ? 'Verifying...' : 'Verify & Sign In' }}
              </button>
              <p class="auth-switch">
                <template v-if="resendCountdown > 0">Resend OTP in <strong>{{ resendCountdown }}s</strong></template>
                <a v-else @click="submitLogin" style="cursor:pointer">Resend OTP</a>
                &nbsp;&middot;&nbsp;
                <a @click="loginStep=0;loginError='';otpError=''" style="cursor:pointer">Change email</a>
              </p>
            </form>
          </template>

          <!-- STEP 2: Role Select -->
          <template v-else-if="loginStep === 2">
            <div class="login-success-badge"><span class="lsb-dot"></span> Authenticated</div>
            <h3 class="modal-title">Welcome, {{ verifiedUser.name }}</h3>
            <p class="modal-sub">Choose your workspace to continue</p>
            <div class="role-cards">
              <a class="role-card" :href="verifiedUser.employerLink || DEFAULT_EMPLOYER_URL" target="_blank">
                <div class="rc-ico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg></div>
                <div class="rc-title">Employer Panel</div>
                <div class="rc-desc">Access your team dashboard, leads, and pipeline operations.</div>
                <div class="rc-arrow">Open Panel →</div>
              </a>
              <a class="role-card" :href="verifiedUser.adminLink || DEFAULT_ADMIN_URL" target="_blank">
                <div class="rc-ico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg></div>
                <div class="rc-title">Admin Panel</div>
                <div class="rc-desc">Manage team configuration, roles, funnels, and enterprise settings.</div>
                <div class="rc-arrow" style="color:#7c3aed">Open Panel →</div>
              </a>
            </div>
            <p class="auth-switch" style="margin-top:14px">Not you? <a @click="logout" style="cursor:pointer;color:#ef4444">Sign out</a></p>
          </template>
        </div>
      </div>
    </Transition>

    <!-- SIGN UP MODAL -->
    <Transition name="modal">
      <div v-if="showSignup" class="modal-overlay" @click.self="showSignup=false">
        <div class="modal-box">
          <div class="modal-hdr">
            <div class="modal-brand"><span class="brand-name sm">nucle<span class="brand-blue">us</span></span></div>
            <button class="modal-close" @click="showSignup=false">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </button>
          </div>
          <h3 class="modal-title">Get started free</h3>
          <p class="modal-sub">14-day free trial · No credit card required</p>
          <div class="signup-steps">
            <div v-for="(st,i) in ['Account','Organization','Done']" :key="st" :class="['ss-step', {active: signupStep>i, current: signupStep===i}]">
              <div class="ss-dot">{{ i < signupStep ? '✓' : i+1 }}</div>
              <span>{{ st }}</span>
            </div>
          </div>
          <div v-if="signupDone" class="auth-success">
            <div class="success-icon">&#10003;</div>
            <h4>Account Created!</h4>
            <p>Your details have been registered. Check your email to get started.</p>
          </div>
          <form v-else @submit.prevent="submitSignup" class="auth-form">
            <div v-if="signupError" class="login-error">{{ signupError }}</div>
            <template v-if="signupStep===0">
              <div class="af-row-2">
                <div class="af-g"><label>First Name</label><input v-model="signup.firstName" placeholder="First name" required @blur="trackHalfFill('signup')"/></div>
                <div class="af-g"><label>Last Name</label><input v-model="signup.lastName" placeholder="Last name" @blur="trackHalfFill('signup')"/></div>
              </div>
              <div class="af-g"><label>Work Email</label><input v-model="signup.email" type="email" placeholder="you@company.com" required @blur="trackHalfFill('signup')"/></div>
              <div class="af-g"><label>Phone</label><input v-model="signup.phone" placeholder="+91 00000 00000" @blur="trackHalfFill('signup')"/></div>
              <div class="af-g"><label>Password</label><input v-model="signup.password" type="password" placeholder="Min. 8 characters" required/></div>
              <button type="button" class="book-btn w100 lg" @click="signupStep=1">Continue</button>
            </template>
            <template v-if="signupStep===1">
              <div class="af-g"><label>Company Name</label><input v-model="signup.company" placeholder="Your company" required @blur="trackHalfFill('signup')"/></div>
              <div class="af-g"><label>Your Role</label>
                <select v-model="signup.role">
                  <option value="">Select your role</option>
                  <option>Founder / CEO</option><option>Sales Manager</option><option>Sales Rep</option><option>RevOps / Operations</option><option>Marketing</option><option>Other</option>
                </select>
              </div>
              <div class="af-g"><label>Team Size</label>
                <select v-model="signup.teamSize">
                  <option value="">Select size</option>
                  <option>1-10</option><option>11-50</option><option>51-200</option><option>200+</option>
                </select>
              </div>
              <div class="af-g"><label>Industry</label>
                <select v-model="signup.industry">
                  <option value="">Select industry</option>
                  <option>Real Estate</option><option>Financial Services</option><option>EdTech</option><option>SaaS / Technology</option><option>Healthcare</option><option>Manufacturing</option><option>Other</option>
                </select>
              </div>
              <button type="submit" class="book-btn w100 lg" :disabled="signupLoading">{{ signupLoading ? 'Creating Account...' : 'Create Account' }}</button>
              <p class="back-link" @click="signupStep=0">← Back</p>
            </template>
            <p class="auth-switch">Already have an account? <a @click="showSignup=false;showLogin=true">Sign in</a></p>
            <p class="auth-terms">By creating an account you agree to our <a href="#">Terms of Service</a> and <a href="#">Privacy Policy</a>.</p>
          </form>
        </div>
      </div>
    </Transition>

    <!-- VIEW PRODUCT MODAL -->
    <Transition name="modal">
      <div v-if="showViewProduct" class="modal-overlay" @click.self="showViewProduct=false">
        <div class="modal-box">
          <div class="modal-hdr">
            <div class="modal-brand"><span class="brand-name sm">nucle<span class="brand-blue">us</span></span></div>
            <button class="modal-close" @click="showViewProduct=false">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </button>
          </div>
          <h3 class="modal-title">Access Live Product</h3>
          <p class="modal-sub">Choose your workspace to continue</p>
          <div class="role-cards">
            <a class="role-card" :href="DEFAULT_EMPLOYER_URL" target="_blank">
              <div class="rc-ico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg></div>
              <div class="rc-title">Employer Panel</div>
              <div class="rc-desc">Team dashboard for lead management, pipeline tracking, and day-to-day operations.</div>
              <div class="rc-arrow">Open Employer Panel →</div>
            </a>
            <a class="role-card" :href="DEFAULT_ADMIN_URL" target="_blank">
              <div class="rc-ico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg></div>
              <div class="rc-title">Admin Panel</div>
              <div class="rc-desc">Full enterprise control — funnels, roles, team configuration, and system-wide analytics.</div>
              <div class="rc-arrow" style="color:#7c3aed">Open Admin Panel →</div>
            </a>
          </div>
        </div>
      </div>
    </Transition>

    <!-- ZIRA FLOATING CHAT WIDGET -->
    <div class="zira-float-wrap">
      <Transition name="zira-panel">
        <div v-if="showZiraChat" class="zira-float-panel">
          <div class="zfp-hdr">
            <div class="zfp-hdr-left">
              <div class="zfp-avatar">Z</div>
              <div>
                <div class="zfp-name">ZIRA — Product AI</div>
                <div class="zfp-status"><span class="zfp-online-dot"></span>Online · Ask anything about Nucleous</div>
              </div>
            </div>
            <button class="zfp-close" @click="showZiraChat=false">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </button>
          </div>
          <div class="zfp-messages" ref="ziraMessagesEl">
            <div v-for="(msg, i) in ziraMessages" :key="i" :class="['zfp-msg', msg.role]">
              <span v-if="msg.role === 'bot'" class="zfp-msg-avatar">Z</span>
              <div class="zfp-msg-bubble">{{ msg.text }}</div>
            </div>
            <div v-if="ziraLoading" class="zfp-msg bot">
              <span class="zfp-msg-avatar">Z</span>
              <div class="zfp-msg-bubble zfp-typing"><span></span><span></span><span></span></div>
            </div>
          </div>
          <div class="zfp-suggestions" v-if="ziraMessages.length === 1">
            <button v-for="s in ziraSuggestions" :key="s" class="zfp-chip" @click="sendZiraMessage(s)">{{ s }}</button>
          </div>
          <div class="zfp-input-row">
            <input v-model="ziraInput" placeholder="Ask about features, pricing, integrations..." @keyup.enter="sendZiraMessage()" :disabled="ziraLoading" ref="ziraInputEl"/>
            <button class="zfp-send" @click="sendZiraMessage()" :disabled="ziraLoading || !ziraInput.trim()">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
            </button>
          </div>
        </div>
      </Transition>

      <button class="zira-float-btn" @click="toggleZiraChat" :class="{open: showZiraChat}">
        <template v-if="!showZiraChat">
          <div class="zfb-content">
            <div class="zfb-avatar-wrap">
              <div class="zfb-av">Z</div>
              <span class="zfb-pulse"></span>
            </div>
            <div class="zfb-text">
              <div class="zfb-label">Ask ZIRA</div>
              <div class="zfb-sub">Product AI · Online</div>
            </div>
          </div>
        </template>
        <template v-else>
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </template>
      </button>
    </div>

  </div>
</template>

<script setup>
import { ref, reactive, onMounted, onUnmounted, nextTick, watch } from 'vue'

// ─── CONFIG ──────────────────────────────────────────────────────────────────
// ⚠️ REPLACE THIS URL with your deployed Apps Script URL after deployment
const SHEET_URL            = 'https://script.google.com/macros/s/AKfycbyfqpoWZ2L2xNw99Z6-90_tFXXpa6IB1oS6gI79ZmgWpMTX7kdzm7rU3oxbOAVW3KU/exec'
const DEFAULT_EMPLOYER_URL = 'https://script.google.com/macros/s/AKfycbynpgWjg4G5-oXJpBcmf0maWHwwH8Z9X92tWvRodRH_-O2TzGZvMLKMZ_nhZyxrFJmxgg/exec'
const DEFAULT_ADMIN_URL    = 'https://script.google.com/macros/s/AKfycbynpgWjg4G5-oXJpBcmf0maWHwwH8Z9X92tWvRodRH_-O2TzGZvMLKMZ_nhZyxrFJmxgg/exec?admin=true'
const SESSION_KEY          = 'nuc_sid'
const LOGIN_KEY            = 'nuc_user'
const LOGIN_TTL_MS         = 7 * 24 * 60 * 60 * 1000

// ─── CORE STATE ──────────────────────────────────────────────────────────────
const scrolled       = ref(false)
const mobOpen        = ref(false)
const activeDropdown = ref(null)
const activeStep     = ref(0)

// ─── MODAL VISIBILITY ────────────────────────────────────────────────────────
const showLogin       = ref(false)
const showSignup      = ref(false)
const showViewProduct = ref(false)

// ─── LOGIN STATE ─────────────────────────────────────────────────────────────
const loginStep       = ref(0)
const loginLoading    = ref(false)
const loginError      = ref('')
const otpValue        = ref('')
const otpLoading      = ref(false)
const otpError        = ref('')
const resendCountdown = ref(0)
let   resendTimer     = null
const verifiedUser    = ref({})
const login           = reactive({ email: '', password: '' })

// ─── SIGNUP STATE ────────────────────────────────────────────────────────────
const signupStep    = ref(0)
const signupLoading = ref(false)
const signupDone    = ref(false)
const signupError   = ref('')
const signup        = reactive({ firstName:'', lastName:'', email:'', password:'', phone:'', company:'', role:'', teamSize:'', industry:'' })

// ─── CONTACT FORM STATE ───────────────────────────────────────────────────────
const contactDone    = ref(false)
const contactSending = ref(false)
const cf             = reactive({ name:'', email:'', company:'', size:'', message:'' })

// ─── TRACKING STATE ───────────────────────────────────────────────────────────
const sessionId         = ref('')
const cfHalfTracked     = ref(false)
const signupHalfTracked = ref(false)
let   halfFillTimeout   = null

// ─── ZIRA CHAT STATE ──────────────────────────────────────────────────────────
const showZiraChat    = ref(false)
const ziraMessages    = ref([
  { role: 'bot', text: "Hi! I'm ZIRA, Nucleous's product AI. I can answer any questions about our platform — features, integrations, how we work, pricing, and more. What would you like to know?" }
])
const ziraHistory     = ref([])
const ziraInput       = ref('')
const ziraLoading     = ref(false)
const ziraMessagesEl  = ref(null)
const ziraInputEl     = ref(null)
const ziraSuggestions = [
  'What does Nucleous do?',
  'How does AI lead scoring work?',
  'What integrations are supported?',
  'Is there a free trial?'
]

// ─── STATIC DATA ─────────────────────────────────────────────────────────────
const trustStats = [
  { num:'50K+', label:'Leads Managed' },
  { num:'40%',  label:'Faster Closure' },
  { num:'94%',  label:'Platform Adoption' },
  { num:'500+', label:'Teams Active' }
]

const kpis = [
  { val:'$248K', label:'Monthly Revenue',  chg:'+12%', color:'#2563eb', up:true },
  { val:'89%',   label:'Closure Rate',     chg:'+8%',  color:'#059669', up:true },
  { val:'342',   label:'Active Leads',     chg:'+24',  color:'#7c3aed', up:true },
  { val:'18',    label:'Deals Closed',     chg:'+5',   color:'#d97706', up:true }
]

const miniBars = [
  { label:'Qualified',   pct:78, color:'#2563eb' },
  { label:'Proposal',    pct:56, color:'#7c3aed' },
  { label:'Negotiation', pct:42, color:'#059669' },
  { label:'Closed',      pct:89, color:'#d97706' }
]

const teamActivity = [
  { initial:'A', name:'Aryan S.',  stat:'14 leads contacted', status:'active', statusLabel:'Active' },
  { initial:'P', name:'Priya M.', stat:'11 deals closed',    status:'top',    statusLabel:'Top'    },
  { initial:'R', name:'Rahul K.', stat:'8 follow-ups done',  status:'active', statusLabel:'Active' }
]

const metricsStrip = [
  { num:'2.4M+',  label:'Pipeline Value Tracked'   },
  { num:'98%',    label:'Uptime SLA'                },
  { num:'4.2min', label:'Avg Follow-up Time'        },
  { num:'3.2x',   label:'Lead Conversion Lift'      },
  { num:'24/7',   label:'AI + Technical Support'    },
  { num:'60+',    label:'Platform Integrations'     }
]

const integrations = ['HubSpot','Salesforce','Pipedrive','WhatsApp','Slack','Zapier','Gmail','Google Sheets','Notion','Zoho']

const aboutPoints = [
  { title:'Unified Lead Lifecycle',    desc:'Every lead tracked from first capture to final closure in one structured system.',              svg:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/></svg>' },
  { title:'AI-Driven Follow-ups',      desc:'ZIRA automates intelligent outreach via email and WhatsApp, adapting tone and timing per lead.', svg:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="4"/></svg>' },
  { title:'Workforce Intelligence',    desc:'Track shifts, productivity metrics, and team performance with real-time accountability.',        svg:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>' },
  { title:'Enterprise Admin Control',  desc:'Role-based permissions, funnel management, and multi-team governance from one admin panel.',    svg:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>' }
]

const aboutStats = [
  { num:'40%',  label:'Faster Lead Closure',  desc:'Vs. manual pipeline management',      color:'#2563eb' },
  { num:'3.2x', label:'Conversion Lift',      desc:'With AI follow-up automation',        color:'#059669' },
  { num:'60%',  label:'Reduced Overhead',     desc:'Operational cost reduction',           color:'#7c3aed' },
  { num:'94%',  label:'Adoption Rate',        desc:'Teams active within 7 days',           color:'#d97706' }
]

const features = [
  { title:'Multi-Source Lead Capture',    desc:'Capture leads from websites, WhatsApp, email, ads, and third-party platforms — all in one inbox.',                          svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M22 16.92v3a2 2 0 01-2.18 2 19.79 19.79 0 01-8.63-3.07A19.5 19.5 0 013.07 9.81a19.79 19.79 0 01-3.07-8.67A2 2 0 012.18 1h3a2 2 0 012 1.72c.127.96.361 1.903.7 2.81a2 2 0 01-.45 2.11L6.09 7.91A16 16 0 0018 19.91l1.27-1.27a2 2 0 012.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0122 16.92z"/></svg>' },
  { title:'AI Lead Scoring',              desc:'Score and qualify every lead by intent, fit, and behavioral signals for prioritized action.',                                svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>' },
  { title:'Funnel Pipeline Management',   desc:'Structured funnels with stage-wise ownership, SLAs, and automated progression triggers.',                                   svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>' },
  { title:'Quotation Generation',         desc:'Build and send accurate, branded quotes from pipeline data — configured once, reused across deals.',                       svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>' },
  { title:'AI Follow-up Automation',      desc:'Scheduled WhatsApp and email follow-ups powered by ZIRA — context-aware, timing-optimized.',                               svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/></svg>' },
  { title:'Workforce Intelligence',       desc:'Track shifts, breaks, activity levels, AHT, and task completion in a unified productivity layer.',                          svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>' },
  { title:'Revenue Forecasting',          desc:'AI-powered quarterly forecasting based on pipeline velocity, team performance, and historical trends.',                     svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>' },
  { title:'Enterprise Admin Control',     desc:'Role-based permissions, funnel governance, lead assignment, and multi-team management from one panel.',                     svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>' },
  { title:'CRM Integrations',             desc:'Native sync with HubSpot, Salesforce, Pipedrive, Zoho, and 50+ tools — zero data silos.',                                 svg:'<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></svg>' }
]

const revLegend    = [{ label:'Won', val:'$1.2M', color:'#2563eb' },{ label:'In Progress', val:'$840K', color:'#059669' },{ label:'Forecasted', val:'$360K', color:'#d97706' }]
const donutData    = [{ label:'Web / Ads', pct:40, color:'#2563eb' },{ label:'WhatsApp', pct:24, color:'#059669' },{ label:'Referral', pct:18, color:'#7c3aed' },{ label:'Other', pct:18, color:'#d97706' }]
const respBars     = [{ day:'M', h:38, active:false },{ day:'T', h:52, active:false },{ day:'W', h:44, active:false },{ day:'T', h:30, active:false },{ day:'F', h:28, active:true },{ day:'S', h:18, active:false },{ day:'S', h:22, active:false }]
const aiPerf       = [{ label:'Open Rate', pct:74, color:'#2563eb' },{ label:'Reply Rate', pct:52, color:'#059669' },{ label:'Conversion', pct:38, color:'#7c3aed' },{ label:'Engagement', pct:81, color:'#d97706' }]
const funnelStages = [{ name:'Captured', count:342, pct:100 },{ name:'Qualified', count:218, pct:64 },{ name:'Proposal', count:134, pct:39 },{ name:'Negotiation', count:87, pct:25 },{ name:'Closed', count:58, pct:17 }]

const steps = [
  { title:'Capture and Qualify',   desc:'Leads enter Nucleous from any channel and are instantly scored, routed, and assigned to the right team member.', points:['Multi-channel lead ingestion','AI scoring on arrival','Automated assignment rules','Instant team notification via WhatsApp'] },
  { title:'Engage and Automate',   desc:'Structured follow-up sequences, ZIRA-guided messaging, and automated pipeline transitions keep every lead progressing.', points:['AI-timed follow-up sequences','ZIRA next-action recommendations','Automatic stage transitions','Quotation generation in-pipeline'] },
  { title:'Close and Analyse',     desc:'Deals close with full audit trails. Analytics surface what is working so leadership can optimize the entire system.', points:['Closure tracking and revenue attribution','AI-powered revenue forecasting','Funnel conversion analytics','Team productivity and performance reporting'] }
]

const solutions = [
  { title:'Real Estate Teams',     sub:'Lead-to-site-visit pipeline management',      svg:'<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>' },
  { title:'Financial Services',    sub:'Regulated lead capture with compliance trails', svg:'<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>' },
  { title:'EdTech / Admissions',   sub:'Student enquiry to enrollment funnels',        svg:'<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>' },
  { title:'SaaS Sales Teams',      sub:'Trial-to-paid pipeline with usage scoring',    svg:'<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8m-4-4v4"/></svg>' }
]

const ziraCaps = [
  { title:'Contextual Lead Guidance',   desc:'ZIRA reads lead history and tells your team exactly what to do next.'                     },
  { title:'Automated Outreach',         desc:'WhatsApp and email campaigns generated, scheduled, and sent without manual effort.'       },
  { title:'Pipeline Intelligence',      desc:'ZIRA flags stale leads, bottlenecks, and at-risk deals before they slip.'                 },
  { title:'Natural Language Queries',   desc:'Ask about your pipeline in plain language and get instant, accurate answers.'             }
]

const contactInfo = [
  { label:'Email',    val:'hello.Nucleous@gmail.com', svg:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>' },
  { label:'Phone',    val:'+919991283530',             svg:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><path d="M22 16.92v3a2 2 0 01-2.18 2 19.79 19.79 0 01-8.63-3.07A19.5 19.5 0 013.07 9.81a19.79 19.79 0 01-3.07-8.67A2 2 0 012.18 1h3a2 2 0 012 1.72c.127.96.361 1.903.7 2.81a2 2 0 01-.45 2.11L6.09 7.91A16 16 0 0018 19.91l1.27-1.27a2 2 0 012.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0122 16.92z"/></svg>' },
  { label:'Location', val:'India',                     svg:'<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>' }
]

const footLinks = [
  { head:'Product',   links:['Features','Integrations','Changelog','Roadmap','Security'] },
  { head:'Solutions', links:['Real Estate','Financial Services','EdTech','SaaS Teams']   },
  { head:'Company',   links:['About','Blog','Careers','Press']                           },
  { head:'Support',   links:['Documentation','API Docs','Contact','Status']              }
]

// ─── WATCHER: restore session from localStorage ───────────────────────────────
watch(showLogin, (val) => {
  if (val) {
    try {
      const stored = localStorage.getItem(LOGIN_KEY)
      if (stored) {
        const u = JSON.parse(stored)
        if (Date.now() - u.loginTime < LOGIN_TTL_MS) {
          verifiedUser.value = u
          loginStep.value    = 2
          return
        }
        localStorage.removeItem(LOGIN_KEY)
      }
    } catch(e) { localStorage.removeItem(LOGIN_KEY) }
    loginStep.value  = 0
    loginError.value = ''
    otpError.value   = ''
    otpValue.value   = ''
  } else {
    clearInterval(resendTimer)
  }
})

// ─── NAVIGATION ──────────────────────────────────────────────────────────────
function goto(id) {
  const el = document.getElementById(id)
  if (el) el.scrollIntoView({ behavior: 'smooth' })
  mobOpen.value        = false
  activeDropdown.value = null
}

function openViewProduct()  { showViewProduct.value = true }

function handleFootLink(lnk) {
  if (lnk === 'Contact')  goto('contact')
  else if (lnk === 'Features') goto('features')
}

function closeLogin() { showLogin.value = false }

// ─── LOGIN: send OTP via Apps Script ─────────────────────────────────────────
async function submitLogin() {
  if (!login.email || !login.password) return
  loginError.value   = ''
  loginLoading.value = true
  clearInterval(resendTimer)
  try {
    const url  = `${SHEET_URL}?action=login&email=${encodeURIComponent(login.email)}&password=${encodeURIComponent(login.password)}`
    const res  = await fetch(url)
    const data = await res.json()
    if (data.success && data.otpSent) {
      loginStep.value = 1
      otpValue.value  = ''
      otpError.value  = ''
      startResendCountdown()
    } else {
      loginError.value = data.message || 'Invalid credentials. Please check your email and password.'
    }
  } catch(e) {
    loginError.value = 'Connection error. Please try again.'
  }
  loginLoading.value = false
}

function startResendCountdown() {
  resendCountdown.value = 30
  resendTimer = setInterval(() => {
    resendCountdown.value--
    if (resendCountdown.value <= 0) clearInterval(resendTimer)
  }, 1000)
}

// ─── OTP VERIFICATION: fetch user + panel URLs from sheet ─────────────────────
async function verifyOTP() {
  if (otpValue.value.length < 6) return
  otpError.value   = ''
  otpLoading.value = true
  try {
    const url  = `${SHEET_URL}?action=verifyOTP&email=${encodeURIComponent(login.email)}&otp=${encodeURIComponent(otpValue.value)}`
    const res  = await fetch(url)
    const data = await res.json()
    if (data.success && data.user) {
      verifiedUser.value = data.user
      localStorage.setItem(LOGIN_KEY, JSON.stringify({ ...data.user, loginTime: Date.now() }))
      loginStep.value = 2
    } else {
      otpError.value = data.message || 'Incorrect or expired OTP. Please try again.'
    }
  } catch(e) {
    otpError.value = 'Verification failed. Please try again.'
  }
  otpLoading.value = false
}

// ─── LOGOUT ───────────────────────────────────────────────────────────────────
function logout() {
  localStorage.removeItem(LOGIN_KEY)
  verifiedUser.value = {}
  loginStep.value    = 0
  loginError.value   = ''
  login.email        = ''
  login.password     = ''
  showLogin.value    = false
}

// ─── SIGN UP ──────────────────────────────────────────────────────────────────
async function submitSignup() {
  if (!signup.firstName || !signup.email) return
  signupError.value   = ''
  signupLoading.value = true
  try {
    const params = new URLSearchParams({
      action:    'register',
      firstName: signup.firstName,
      lastName:  signup.lastName,
      email:     signup.email,
      phone:     signup.phone,
      password:  signup.password,
      company:   signup.company,
      role:      signup.role,
      teamSize:  signup.teamSize,
      industry:  signup.industry
    })
    const res  = await fetch(`${SHEET_URL}?${params}`)
    const data = await res.json()
    if (!data.success) {
      signupError.value = data.message || 'Registration failed. Please try again.'
      signupLoading.value = false
      return
    }
  } catch(e) { /* proceed */ }
  signupLoading.value = false
  signupDone.value    = true
  signupStep.value    = 2
}

// ─── CONTACT FORM ─────────────────────────────────────────────────────────────
async function submitContact() {
  if (!cf.name || !cf.email) return
  contactSending.value = true
  try {
    const params = new URLSearchParams({
      action:   'contact',
      name:     cf.name,
      email:    cf.email,
      company:  cf.company,
      size:     cf.size,
      message:  cf.message
    })
    await fetch(`${SHEET_URL}?${params}`)
  } catch(e) {}
  contactSending.value = false
  contactDone.value    = true
  cfHalfTracked.value  = false
}

// ─── VISITOR TRACKING ─────────────────────────────────────────────────────────
function trackVisitor(sid) {
  try {
    const params = new URLSearchParams({
      action:    'trackVisitor',
      sessionId: sid,
      userAgent: navigator.userAgent.substring(0, 200),
      referrer:  document.referrer.substring(0, 200),
      page:      window.location.href.substring(0, 200)
    })
    fetch(`${SHEET_URL}?${params}`)
  } catch(e) {}
}

// ─── HALF-FILL TRACKING ───────────────────────────────────────────────────────
function trackHalfFill(formType) {
  clearTimeout(halfFillTimeout)
  halfFillTimeout = setTimeout(() => {
    let hasContent = false
    let data = {}
    if (formType === 'contact') {
      data = { name: cf.name, email: cf.email, company: cf.company, message: cf.message }
      hasContent = !!(cf.name || cf.email || cf.company || cf.message)
      if (!hasContent || cfHalfTracked.value) return
      cfHalfTracked.value = true
    } else if (formType === 'signup') {
      data = { firstName: signup.firstName, lastName: signup.lastName, email: signup.email, company: signup.company }
      hasContent = !!(signup.firstName || signup.email)
      if (!hasContent || signupHalfTracked.value) return
      signupHalfTracked.value = true
    }
    try {
      const params = new URLSearchParams({ action: 'trackHalfFill', sessionId: sessionId.value, formType, data: JSON.stringify(data) })
      fetch(`${SHEET_URL}?${params}`)
    } catch(e) {}
  }, 2500)
}

// ─── ZIRA CHAT ────────────────────────────────────────────────────────────────
function toggleZiraChat() {
  showZiraChat.value = !showZiraChat.value
  if (showZiraChat.value) {
    nextTick(() => {
      scrollZiraToBottom()
      if (ziraInputEl.value) ziraInputEl.value.focus()
    })
  }
}

async function sendZiraMessage(text) {
  const msg = (text || ziraInput.value || '').trim()
  if (!msg || ziraLoading.value) return
  ziraInput.value = ''
  ziraMessages.value.push({ role: 'user', text: msg })
  ziraHistory.value.push({ role: 'user', content: msg })
  ziraLoading.value = true
  await nextTick()
  scrollZiraToBottom()
  try {
    const params = new URLSearchParams({ action: 'chat', messages: JSON.stringify(ziraHistory.value) })
    const res  = await fetch(`${SHEET_URL}?${params}`)
    const data = await res.json()
    const reply = (data.success && data.reply) ? data.reply : "I'm having a moment — please try again or reach us at hello.Nucleous@gmail.com"
    ziraMessages.value.push({ role: 'bot', text: reply })
    ziraHistory.value.push({ role: 'assistant', content: reply })
  } catch(e) {
    ziraMessages.value.push({ role: 'bot', text: 'Connection issue. Reach us at hello.Nucleous@gmail.com or +919991283530.' })
  }
  ziraLoading.value = false
  await nextTick()
  scrollZiraToBottom()
}

function scrollZiraToBottom() {
  if (ziraMessagesEl.value) ziraMessagesEl.value.scrollTop = ziraMessagesEl.value.scrollHeight
}

// ─── SCROLL ───────────────────────────────────────────────────────────────────
function onScroll() { scrolled.value = window.scrollY > 48 }

// ─── LIFECYCLE ───────────────────────────────────────────────────────────────
onMounted(() => {
  window.addEventListener('scroll', onScroll)
  let sid = sessionStorage.getItem(SESSION_KEY)
  if (!sid) {
    sid = 'sid_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9)
    sessionStorage.setItem(SESSION_KEY, sid)
    trackVisitor(sid)
  }
  sessionId.value = sid
})

onUnmounted(() => {
  window.removeEventListener('scroll', onScroll)
  clearInterval(resendTimer)
  clearTimeout(halfFillTimeout)
})
</script>

<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');

:root {
  --bg:#ffffff; --bg2:#f7f8fc; --bg3:#eff1f8; --bg4:#e6e9f4;
  --text:#0c0d1a; --text2:#4a5270; --text3:#8892b0; --text4:#b8c0d4;
  --blue:#2563eb; --blue2:#1d4ed8; --blue-light:#eff6ff; --blue-mid:#dbeafe; --blue-border:rgba(37,99,235,0.15);
  --border:#e4e8f4; --border2:#edf0f9;
  --green:#059669; --green-light:#ecfdf5; --purple:#7c3aed; --amber:#d97706;
  --sh-sm:0 1px 4px rgba(0,0,0,.04),0 2px 8px rgba(0,0,0,.03);
  --sh:0 2px 8px rgba(0,0,0,.05),0 4px 20px rgba(0,0,0,.03);
  --sh-md:0 4px 16px rgba(0,0,0,.07),0 8px 32px rgba(0,0,0,.04);
  --sh-lg:0 8px 32px rgba(0,0,0,.09),0 20px 60px rgba(0,0,0,.06);
  --sh-xl:0 16px 48px rgba(0,0,0,.10),0 32px 80px rgba(0,0,0,.07);
  --sh-blue:0 4px 20px rgba(37,99,235,.25);
  --r:9px; --r-md:13px; --r-lg:17px; --r-xl:21px; --r-2xl:26px;
  --fn:'DM Sans',sans-serif;
  --ease:.18s cubic-bezier(.4,0,.2,1);
  touch-action:pan-x pan-y;
}

*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}
html{scroll-behavior:smooth;-webkit-text-size-adjust:100%;touch-action:manipulation}
body{font-family:var(--fn);background:var(--bg);color:var(--text);overflow-x:hidden;line-height:1.6;-webkit-font-smoothing:antialiased;font-size:13px;touch-action:manipulation}
input,select,textarea,button{font-family:var(--fn);font-size:13px;-webkit-appearance:none;appearance:none;background:none;border:none;outline:none;touch-action:manipulation;-webkit-text-size-adjust:none;text-size-adjust:none}
a{color:inherit;text-decoration:none;cursor:pointer}
img{display:block;max-width:100%}
ul{list-style:none}

input,select,textarea{font-size:13px!important}
@media screen and (max-width:768px){input,select,textarea{font-size:16px!important}}

.blue-text{color:var(--blue)}
.brand-blue{color:var(--blue)}
.container{max-width:1140px;margin:0 auto;padding:0 32px}
.section{padding:88px 0}
.gray-bg{background:var(--bg2)}
.section-center{text-align:center;margin-bottom:52px}
.centered-text{max-width:480px;margin:0 auto}
.w100{width:100%;display:block;text-align:center}
.w50{width:calc(50% - 5px);text-align:center;justify-content:center}

.eyebrow{display:inline-flex;align-items:center;gap:7px;font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--blue);background:var(--blue-light);border:1px solid var(--blue-border);padding:5px 12px;border-radius:100px;margin-bottom:12px}
.section-title{font-size:clamp(22px,3.2vw,40px);font-weight:800;line-height:1.14;letter-spacing:-.03em;margin-bottom:14px;color:var(--text)}
.body-text{font-size:14px;color:var(--text2);line-height:1.8}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:60px;align-items:center}

/* BUTTONS */
.book-btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;background:var(--blue);color:#fff;border:none;padding:10px 22px;border-radius:var(--r);font-size:13px;font-weight:600;cursor:pointer;transition:all var(--ease);box-shadow:var(--sh-blue);white-space:nowrap;letter-spacing:-.01em}
.book-btn:hover:not(:disabled){background:var(--blue2);transform:translateY(-1px);box-shadow:0 6px 24px rgba(37,99,235,.35)}
.book-btn:disabled{opacity:.55;cursor:default}
.book-btn.lg{padding:12px 26px;font-size:14px;border-radius:var(--r-md)}
.book-btn.sm{padding:7px 15px;font-size:12px}
.book-btn.w100{width:100%}
.ghost-btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;background:transparent;color:var(--text);border:1.5px solid var(--border);padding:10px 22px;border-radius:var(--r);font-size:13px;font-weight:500;cursor:pointer;transition:all var(--ease);white-space:nowrap}
.ghost-btn:hover{border-color:var(--blue);color:var(--blue);background:var(--blue-light)}
.ghost-btn.lg{padding:12px 26px;font-size:14px;border-radius:var(--r-md)}
.view-product-btn{display:inline-flex;align-items:center;gap:5px;background:transparent;color:var(--blue);border:1.5px solid var(--blue-border);padding:8px 16px;border-radius:var(--r);font-size:13px;font-weight:600;cursor:pointer;transition:all var(--ease)}
.view-product-btn:hover{background:var(--blue-light);border-color:var(--blue)}

/* HEADER */
header{position:fixed;top:0;left:0;right:0;z-index:200;background:rgba(255,255,255,.93);backdrop-filter:blur(20px);-webkit-backdrop-filter:blur(20px);border-bottom:1px solid transparent;transition:all var(--ease)}
header.scrolled{border-bottom-color:var(--border);box-shadow:0 1px 0 var(--border),0 2px 16px rgba(0,0,0,.04)}
.hdr{max-width:1140px;margin:0 auto;padding:0 32px;height:62px;display:flex;align-items:center;justify-content:space-between;gap:16px}
.brand{font-size:18px;font-weight:800;display:flex;align-items:center;cursor:pointer;color:var(--text);letter-spacing:-.03em;flex-shrink:0}
.brand-name{font-size:18px;font-weight:800;letter-spacing:-.03em}
.brand-name.sm{font-size:16px}

.nav-links{display:flex;gap:0;flex:1;justify-content:center}
.nav-item{position:relative}
.nav-link{display:flex;align-items:center;gap:4px;font-size:13px;font-weight:500;color:var(--text2);padding:7px 13px;border-radius:8px;transition:all var(--ease);cursor:pointer;white-space:nowrap}
.nav-link:hover{color:var(--text);background:var(--bg2)}
a.nav-link{font-size:13px;font-weight:500;color:var(--text2);padding:7px 13px;border-radius:8px;transition:all var(--ease);cursor:pointer;white-space:nowrap}
a.nav-link:hover{color:var(--text);background:var(--bg2)}

.dropdown{position:absolute;top:calc(100% + 8px);left:50%;transform:translateX(-50%);background:#fff;border:1px solid var(--border);border-radius:var(--r-lg);box-shadow:var(--sh-lg);min-width:420px;display:flex;gap:0;z-index:300;overflow:hidden}
.dd-section{flex:1;padding:16px;border-right:1px solid var(--border2)}
.dd-section:last-child{border-right:none}
.dd-head{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text3);padding:4px 8px 8px;margin-bottom:4px}
.dd-item{display:flex;align-items:flex-start;gap:10px;padding:9px 8px;border-radius:var(--r);cursor:pointer;transition:background var(--ease)}
.dd-item:hover{background:var(--bg2)}
.dd-ico{width:30px;height:30px;min-width:30px;border-radius:8px;background:var(--bg3);display:flex;align-items:center;justify-content:center;color:var(--text2)}
.dd-title{font-size:13px;font-weight:600;color:var(--text);margin-bottom:2px}
.dd-sub{font-size:11px;color:var(--text3);line-height:1.4}
.hdr-right{display:flex;align-items:center;gap:8px;flex-shrink:0}

.mob-menu{display:none;flex-direction:column;gap:5px;padding:7px;cursor:pointer;border-radius:8px}
.mob-menu span{display:block;width:18px;height:2px;background:var(--text);border-radius:2px;transition:all .22s ease;transform-origin:center}
.mob-menu span.o1{transform:rotate(45deg) translate(4px,4px)}
.mob-menu span.o2{opacity:0;transform:scaleX(0)}
.mob-menu span.o3{transform:rotate(-45deg) translate(4px,-4px)}
.mob-nav{position:fixed;top:62px;left:0;right:0;z-index:199;background:#fff;border-bottom:1px solid var(--border);padding:12px 32px 24px;display:flex;flex-direction:column;gap:2px;box-shadow:var(--sh-md)}
.mob-nav a{font-size:14px;font-weight:500;color:var(--text);padding:12px 0;border-bottom:1px solid var(--border2)}
.mob-auth-row{display:flex;gap:10px;margin-top:14px}
.mob-drop-enter-active,.mob-drop-leave-active{transition:all .18s ease}
.mob-drop-enter-from,.mob-drop-leave-to{opacity:0;transform:translateY(-6px)}

/* HERO */
.hero-section{min-height:100vh;display:flex;flex-direction:column;align-items:center;padding-top:62px;position:relative;overflow:hidden;background:var(--bg)}
.hero-bg{position:absolute;inset:0;z-index:0;overflow:hidden}
.hero-grid-bg{position:absolute;inset:0;background-image:linear-gradient(var(--border2) 1px,transparent 1px),linear-gradient(90deg,var(--border2) 1px,transparent 1px);background-size:44px 44px;opacity:.6}
.hero-overlay{position:absolute;inset:0;background:radial-gradient(ellipse 70% 50% at 50% 0%,rgba(37,99,235,.07) 0%,transparent 65%),linear-gradient(180deg,rgba(255,255,255,0) 0%,rgba(255,255,255,.7) 60%,rgba(255,255,255,1) 100%)}
.hero-content{position:relative;z-index:2;display:flex;flex-direction:column;align-items:center;text-align:center;padding:72px 0 28px}
.hero-badge{display:inline-flex;align-items:center;gap:8px;background:#fff;border:1px solid var(--border);border-radius:100px;padding:6px 16px 6px 10px;font-size:12px;font-weight:500;color:var(--text2);margin-bottom:24px;box-shadow:var(--sh-sm)}
.badge-dot{width:7px;height:7px;border-radius:50%;background:var(--green);box-shadow:0 0 0 3px rgba(5,150,105,.15);flex-shrink:0;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{box-shadow:0 0 0 3px rgba(5,150,105,.15)}50%{box-shadow:0 0 0 5px rgba(5,150,105,.08)}}
.hero-title{font-size:clamp(34px,5.2vw,64px);font-weight:800;line-height:1.08;letter-spacing:-.04em;color:var(--text);max-width:780px;margin-bottom:20px;animation:fadeUp .45s ease .1s both}
.hero-sub{font-size:clamp(13px,1.5vw,16px);color:var(--text2);max-width:500px;line-height:1.8;margin-bottom:30px;animation:fadeUp .45s ease .2s both}
.hero-btns{display:flex;gap:12px;flex-wrap:wrap;justify-content:center;margin-bottom:36px;animation:fadeUp .45s ease .3s both}
.hero-trust-row{display:flex;gap:0;background:#fff;border:1px solid var(--border);border-radius:var(--r-lg);overflow:hidden;box-shadow:var(--sh-sm);animation:fadeUp .45s ease .4s both}
.ht-item{padding:14px 26px;border-right:1px solid var(--border);text-align:center}
.ht-item:last-child{border-right:none}
.ht-num{display:block;font-size:20px;font-weight:800;color:var(--text);letter-spacing:-.03em;line-height:1}
.ht-lbl{display:block;font-size:11px;color:var(--text3);margin-top:3px;font-weight:500}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}

.hero-dash{position:relative;z-index:2;margin-top:40px;animation:fadeUp .55s ease .5s both}
.dash-window{background:#fff;border:1px solid var(--border);border-radius:var(--r-xl);overflow:hidden;box-shadow:var(--sh-xl)}
.dash-topbar{height:36px;background:var(--bg2);display:flex;align-items:center;gap:6px;padding:0 16px;border-bottom:1px solid var(--border)}
.dtb-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.dtb-dot.r{background:#ff5f57}.dtb-dot.y{background:#ffbd2e}.dtb-dot.g{background:#28c840}
.dtb-title{font-size:11px;font-weight:600;color:var(--text2);margin-left:8px;flex:1}
.dtb-live{font-size:10px;font-weight:700;color:var(--green);background:var(--green-light);padding:2px 9px;border-radius:100px;border:1px solid rgba(5,150,105,.2)}
.dash-body{display:grid;grid-template-columns:140px 1fr}
.dash-side{border-right:1px solid var(--border);padding:8px 0}
.dsi{display:flex;align-items:center;gap:7px;font-size:12px;font-weight:500;color:var(--text2);padding:7px 14px;cursor:pointer;transition:all var(--ease)}
.dsi.active{color:var(--blue);background:var(--blue-light);font-weight:600;border-right:2px solid var(--blue)}
.dsi:hover:not(.active){background:var(--bg2)}
.dash-main{padding:14px;display:flex;flex-direction:column;gap:12px}
.dash-kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}
.kpi{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r);padding:11px;text-align:center}
.kpi-val{font-size:16px;font-weight:800;line-height:1;letter-spacing:-.02em}
.kpi-label{font-size:10px;color:var(--text3);margin:3px 0 2px;font-weight:500}
.kpi-chg{font-size:10px;font-weight:600}
.kpi-chg.up{color:var(--green)}.kpi-chg.dn{color:#ef4444}
.dash-charts-row{display:grid;grid-template-columns:2fr 1.5fr 1.5fr;gap:8px}
.dash-chart-box{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r);padding:10px}
.dcb-label{font-size:10px;font-weight:600;color:var(--text3);margin-bottom:8px;text-transform:uppercase;letter-spacing:.06em}
.mini-bars{display:flex;flex-direction:column;gap:6px}
.mb-item{display:grid;grid-template-columns:1fr auto;align-items:center;gap:6px}
.mb-bar-wrap{grid-column:1;background:var(--bg3);border-radius:2px;height:5px;overflow:hidden}
.mb-bar{height:100%;border-radius:2px;transition:width .6s ease}
.mb-label{grid-column:1;font-size:9px;color:var(--text3);margin-top:-2px}
.mb-val{grid-row:1/3;font-size:10px;font-weight:700;color:var(--text)}
.team-activity{display:flex;flex-direction:column;gap:6px}
.ta-row{display:flex;align-items:center;gap:8px}
.ta-avatar{width:24px;height:24px;border-radius:50%;background:var(--blue);color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.ta-info{flex:1;min-width:0}
.ta-name{font-size:11px;font-weight:600;color:var(--text)}
.ta-stat{font-size:10px;color:var(--text3)}
.ta-badge{font-size:9px;font-weight:700;padding:2px 7px;border-radius:100px}
.ta-badge.active{background:#dbeafe;color:var(--blue)}
.ta-badge.top{background:#fef9c3;color:#b45309}

/* METRICS STRIP */
.metrics-strip{background:var(--text);padding:14px 0}
.ms-inner{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}
.ms-item{text-align:center;flex:1;min-width:100px}
.ms-num{font-size:18px;font-weight:800;color:#fff;letter-spacing:-.02em;line-height:1}
.ms-label{font-size:10px;color:rgba(255,255,255,.45);margin-top:2px;font-weight:500;text-transform:uppercase;letter-spacing:.06em}

/* LOGOS */
.logos-strip{background:var(--bg);border-top:1px solid var(--border);border-bottom:1px solid var(--border);padding:14px 0}
.ls-inner{display:flex;align-items:center;gap:20px;flex-wrap:wrap}
.ls-label{font-size:11px;font-weight:600;color:var(--text3);letter-spacing:.08em;text-transform:uppercase;white-space:nowrap;flex-shrink:0}
.ls-logos{display:flex;gap:6px;flex-wrap:wrap}
.ls-logo{font-size:12px;font-weight:600;color:var(--text3);background:var(--bg2);border:1px solid var(--border);padding:4px 12px;border-radius:100px;transition:all var(--ease)}

/* ABOUT */
.about-points{display:flex;flex-direction:column;gap:8px;margin:20px 0 24px}
.apt{display:flex;align-items:flex-start;gap:12px;padding:12px;border-radius:var(--r);border:1px solid var(--border);background:#fff;transition:all var(--ease)}
.apt:hover{border-color:var(--blue-border);background:var(--blue-light);transform:translateX(3px)}
.apt-icon{width:30px;height:30px;min-width:30px;border-radius:8px;background:var(--blue-light);border:1px solid var(--blue-border);display:flex;align-items:center;justify-content:center}
.apt-title{font-size:13px;font-weight:700;margin-bottom:2px;color:var(--text)}
.apt-desc{font-size:12px;color:var(--text2);line-height:1.5}
.about-visual{display:flex;flex-direction:column;gap:12px}
.av-stats-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.avs-card{background:#fff;border:1px solid var(--border);border-top:3px solid;border-radius:var(--r-md);padding:18px}
.avs-num{font-size:28px;font-weight:800;letter-spacing:-.03em;line-height:1}
.avs-label{font-size:12px;font-weight:700;color:var(--text);margin:4px 0 2px}
.avs-desc{font-size:11px;color:var(--text3)}
.av-card.dark-card{background:linear-gradient(135deg,#060d28,#0f1f60);border-radius:var(--r-lg);padding:24px}
.dark-card .avc-label{font-size:13px;font-weight:700;color:#fff;margin-bottom:6px;letter-spacing:-.01em}
.dark-card .avc-desc{font-size:12px;color:rgba(255,255,255,.55);line-height:1.65;margin-bottom:16px}
.avc-btn{background:rgba(255,255,255,.12);color:#fff;border:1px solid rgba(255,255,255,.2);padding:8px 18px;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;transition:all var(--ease)}
.avc-btn:hover{background:rgba(255,255,255,.22)}

/* FEATURES */
.feat-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
.feat-card{background:#fff;border:1px solid var(--border);border-radius:var(--r-lg);padding:22px;transition:all var(--ease)}
.feat-card:hover{transform:translateY(-3px);box-shadow:var(--sh-md);border-color:var(--blue-border)}
.fc-icon{width:40px;height:40px;border-radius:var(--r);background:var(--blue-light);border:1px solid var(--blue-border);display:flex;align-items:center;justify-content:center;color:var(--blue);margin-bottom:12px}
.fc-title{font-size:13px;font-weight:700;margin-bottom:6px;color:var(--text);letter-spacing:-.01em}
.fc-desc{font-size:12px;color:var(--text2);line-height:1.6}

/* ANALYTICS */
.analytics-section{background:#fff}
.analytics-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}
.an-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r-lg);padding:20px}
.an-card.large{grid-column:span 2}
.an-card-hdr{margin-bottom:10px}
.an-card-title{font-size:13px;font-weight:700;color:var(--text);letter-spacing:-.01em}
.an-card-sub{font-size:11px;color:var(--text3);margin-top:2px}
.an-big-num{font-size:28px;font-weight:800;letter-spacing:-.03em;color:var(--text);line-height:1;display:flex;align-items:center;gap:10px}
.an-big-num.blue{color:var(--blue)}
.an-chg{font-size:12px;font-weight:700}
.an-chg.up{color:var(--green)}.an-chg.dn{color:#ef4444}
.an-legend-row{display:flex;gap:16px;flex-wrap:wrap;margin-top:10px}
.an-leg{font-size:11px;color:var(--text2);display:flex;align-items:center;gap:5px}
.al-dot{width:8px;height:8px;border-radius:50%;display:inline-block;flex-shrink:0}
.donut-wrap{display:flex;align-items:center;gap:20px;margin-top:8px}
.donut-legend{display:flex;flex-direction:column;gap:7px}
.dl-item{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text2)}
.dl-item strong{margin-left:auto;color:var(--text);font-weight:700}
.resp-bars{display:flex;align-items:flex-end;gap:6px;height:60px;margin-top:12px}
.rb-item{flex:1;display:flex;flex-direction:column;align-items:center;gap:4px}
.rb-bar-wrap{flex:1;display:flex;align-items:flex-end;width:100%}
.rb-bar{width:100%;border-radius:3px 3px 0 0;transition:height .4s ease}
.rb-label{font-size:9px;color:var(--text3);font-weight:500}
.ai-perf-rows{display:flex;flex-direction:column;gap:10px;margin-top:8px}
.ap-row{display:flex;align-items:center;gap:8px}
.ap-label{font-size:11px;color:var(--text2);width:76px;flex-shrink:0}
.ap-bar-wrap{flex:1;background:var(--bg3);border-radius:2px;height:6px;overflow:hidden}
.ap-bar{height:100%;border-radius:2px}
.ap-val{font-size:11px;font-weight:700;color:var(--text);width:30px;text-align:right}
.funnel-stages{display:flex;flex-direction:column;gap:8px;margin-top:8px}
.fs-row{display:flex;align-items:center;gap:8px}
.fs-name{font-size:11px;color:var(--text2);width:80px;flex-shrink:0}
.fs-bar-wrap{flex:1;background:var(--bg3);border-radius:2px;height:8px;overflow:hidden}
.fs-bar{height:100%;background:linear-gradient(90deg,#2563eb,#7c3aed);border-radius:2px}
.fs-count{font-size:11px;font-weight:700;color:var(--text);width:28px;text-align:right}

/* HOW IT WORKS */
.steps-row{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
.step-card{background:#fff;border:1.5px solid var(--border);border-radius:var(--r-xl);padding:26px 20px;cursor:pointer;transition:all var(--ease)}
.step-card:hover,.step-card.active{border-color:var(--blue-border);box-shadow:var(--sh-md);transform:translateY(-3px)}
.step-card.active{background:var(--blue-light)}
.sc-num{font-size:22px;font-weight:800;color:var(--blue-border);letter-spacing:-.03em;margin-bottom:12px}
.step-card.active .sc-num{color:var(--blue)}
.sc-title{font-size:15px;font-weight:700;margin-bottom:8px;color:var(--text);letter-spacing:-.01em}
.sc-desc{font-size:12px;color:var(--text2);line-height:1.65}
.sc-list{margin-top:14px;display:flex;flex-direction:column}
.sc-list li{font-size:12px;color:var(--text2);padding:6px 0 6px 16px;position:relative;border-bottom:1px solid var(--border2);line-height:1.5}
.sc-list li::before{content:'';position:absolute;left:0;top:50%;transform:translateY(-50%);width:5px;height:5px;border-radius:50%;background:var(--blue)}

/* ZIRA SECTION */
.zira-section{background:var(--bg)}
.zira-chat-mock{background:#fff;border:1px solid var(--border);border-radius:var(--r-xl);overflow:hidden;box-shadow:var(--sh-lg)}
.zcm-hdr{display:flex;align-items:center;gap:10px;padding:16px;background:linear-gradient(135deg,#060d28,#0f1f60);border-bottom:1px solid rgba(255,255,255,.1)}
.zcm-avatar{width:34px;height:34px;border-radius:10px;background:var(--blue);color:#fff;font-weight:800;font-size:15px;display:flex;align-items:center;justify-content:center}
.zcm-name{font-size:13px;font-weight:700;color:#fff}
.zcm-status{font-size:11px;color:rgba(255,255,255,.45)}
.zcm-messages{display:flex;flex-direction:column;gap:10px;padding:16px}
.zcm-msg{max-width:85%;padding:10px 14px;border-radius:12px;font-size:12px;line-height:1.6}
.zcm-msg.bot{background:var(--blue-light);border:1px solid var(--blue-border);color:var(--text);align-self:flex-start;border-radius:4px 12px 12px 12px}
.zcm-msg.user{background:var(--blue);color:#fff;align-self:flex-end;border-radius:12px 12px 4px 12px}
.zcm-input{padding:12px 16px;font-size:12px;color:var(--text3);background:var(--bg2);border-top:1px solid var(--border)}
.zira-caps{display:flex;flex-direction:column;gap:10px;margin:20px 0 24px}
.zc-item{display:flex;align-items:flex-start;gap:12px}
.zc-dot{width:8px;height:8px;min-width:8px;border-radius:50%;background:var(--blue);margin-top:5px}
.zc-title{font-size:13px;font-weight:700;color:var(--text);margin-bottom:2px}
.zc-desc{font-size:12px;color:var(--text2);line-height:1.55}

/* CONTACT */
.contact-info{display:flex;flex-direction:column;gap:14px;margin-top:24px}
.ci-row{display:flex;align-items:flex-start;gap:12px}
.ci-ico{width:38px;height:38px;border-radius:var(--r);background:var(--blue-light);border:1px solid var(--blue-border);display:flex;align-items:center;justify-content:center;flex-shrink:0}
.ci-lbl{font-size:10px;color:var(--text3);font-weight:600;text-transform:uppercase;letter-spacing:.08em;margin-bottom:1px}
.ci-val{font-size:13px;font-weight:600;color:var(--text)}
.contact-form-wrap{background:#fff;border:1px solid var(--border);border-radius:var(--r-xl);padding:28px;box-shadow:var(--sh-md)}
.contact-form{display:flex;flex-direction:column}
.cf-row-2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.cf-g{margin-bottom:14px}
.cf-g label{display:block;font-size:12px;font-weight:600;color:var(--text2);margin-bottom:5px;letter-spacing:-.01em}
.cf-g input,.cf-g select,.cf-g textarea{width:100%;background:var(--bg2);border:1.5px solid var(--border);color:var(--text);padding:9px 12px;border-radius:var(--r);font-size:13px;transition:border-color .18s,box-shadow .18s;resize:vertical;line-height:1.5}
.cf-g input:focus,.cf-g select:focus,.cf-g textarea:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.08);background:#fff}
.cf-g input::placeholder,.cf-g textarea::placeholder{color:var(--text4)}
.cf-g select option{background:#fff;color:var(--text)}
.form-success{text-align:center;padding:36px 16px}
.form-success h4{font-size:17px;font-weight:700;margin:12px 0 6px;color:var(--text)}
.form-success p{font-size:13px;color:var(--text2)}
.success-icon{width:52px;height:52px;border-radius:50%;background:var(--green-light);border:2px solid var(--green);color:var(--green);font-size:22px;font-weight:700;display:flex;align-items:center;justify-content:center;margin:0 auto}

/* CTA */
.cta-section{background:linear-gradient(135deg,#060d1e 0%,#0f1f60 50%,#060d1e 100%);padding:76px 0;position:relative;overflow:hidden}
.cta-section::before{content:'';position:absolute;top:-80px;left:50%;transform:translateX(-50%);width:500px;height:350px;background:radial-gradient(ellipse,rgba(37,99,235,.18) 0%,transparent 65%);pointer-events:none}
.cta-inner{display:flex;align-items:center;justify-content:space-between;gap:40px;flex-wrap:wrap;position:relative;z-index:1}
.cta-left{flex:1;min-width:260px}
.cta-title{font-size:clamp(20px,2.8vw,34px);font-weight:800;color:#fff;line-height:1.2;margin-bottom:10px;letter-spacing:-.03em}
.cta-sub{font-size:14px;color:rgba(255,255,255,.5);line-height:1.75;max-width:420px}
.cta-right{display:flex;flex-direction:column;align-items:flex-end;gap:10px;flex-shrink:0}
.cta-btn-w{display:inline-flex;align-items:center;justify-content:center;background:#fff;color:var(--blue);border:none;padding:13px 28px;border-radius:var(--r-md);font-size:13px;font-weight:700;cursor:pointer;transition:all var(--ease);box-shadow:0 4px 20px rgba(0,0,0,.2);white-space:nowrap}
.cta-btn-w:hover{transform:translateY(-2px);box-shadow:0 8px 28px rgba(0,0,0,.28)}
.cta-btn-o{display:inline-flex;align-items:center;justify-content:center;background:rgba(255,255,255,.08);color:#fff;border:1.5px solid rgba(255,255,255,.2);padding:11px 26px;border-radius:var(--r-md);font-size:13px;font-weight:500;cursor:pointer;transition:all var(--ease);white-space:nowrap}
.cta-btn-o:hover{background:rgba(255,255,255,.14)}
.cta-note{font-size:11px;color:rgba(255,255,255,.3);text-align:right}

/* FOOTER */
footer{background:#fff;border-top:1px solid var(--border);padding:52px 0 0}
.foot-grid{display:grid;grid-template-columns:1.4fr 2fr;gap:56px;padding-bottom:40px}
.foot-desc{font-size:12px;color:var(--text3);margin-top:10px;line-height:1.65;max-width:260px}
.foot-socials{display:flex;gap:7px;margin-top:16px}
.soc-lnk{width:30px;height:30px;border-radius:8px;border:1px solid var(--border);background:var(--bg2);color:var(--text2);font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all var(--ease)}
.soc-lnk:hover{border-color:var(--blue);color:var(--blue);background:var(--blue-light)}
.foot-links-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:20px}
.fl-col{display:flex;flex-direction:column;gap:8px}
.fl-head{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--text3);margin-bottom:4px}
.fl-col a{font-size:12px;font-weight:500;color:var(--text2);transition:color var(--ease);cursor:pointer}
.fl-col a:hover{color:var(--blue)}
.foot-bottom{border-top:1px solid var(--border);padding:16px 0}
.fbt-inner{display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;font-size:11px;color:var(--text3)}

/* MODAL */
.modal-overlay{position:fixed;inset:0;z-index:999;background:rgba(0,0,0,.5);backdrop-filter:blur(10px);display:flex;align-items:center;justify-content:center;padding:20px}
.modal-box{background:#fff;border:1px solid var(--border);border-radius:var(--r-2xl);width:100%;max-width:440px;max-height:92vh;overflow-y:auto;padding:26px;box-shadow:var(--sh-xl);scrollbar-width:thin}
.modal-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px}
.modal-brand{font-size:15px;font-weight:800}
.modal-close{width:30px;height:30px;border-radius:50%;border:1px solid var(--border);background:var(--bg2);color:var(--text2);display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all var(--ease);flex-shrink:0}
.modal-close:hover{background:var(--bg3);color:var(--text)}
.modal-title{font-size:20px;font-weight:800;letter-spacing:-.025em;color:var(--text);margin-bottom:5px}
.modal-sub{font-size:13px;color:var(--text2);margin-bottom:22px}

/* LOGIN OTP STEP */
.otp-header-icon{width:52px;height:52px;border-radius:var(--r-lg);background:var(--blue-light);border:1px solid var(--blue-border);display:flex;align-items:center;justify-content:center;margin-bottom:14px}
.otp-g{position:relative}
.otp-input{width:100%;background:var(--bg2);border:2px solid var(--border);color:var(--text);padding:14px 18px;border-radius:var(--r-md);font-size:28px!important;font-weight:800;letter-spacing:10px;text-align:center;transition:border-color .18s,box-shadow .18s}
.otp-input:focus{border-color:var(--blue);box-shadow:0 0 0 4px rgba(37,99,235,.1);background:#fff}
.otp-input::placeholder{font-size:22px;letter-spacing:8px;color:var(--text4);font-weight:400}
.otp-hint{font-size:11px;color:var(--text3);text-align:center;margin-top:7px}

/* LOGIN SUCCESS / ROLE SELECT */
.login-success-badge{display:inline-flex;align-items:center;gap:6px;background:var(--green-light);border:1px solid rgba(5,150,105,.2);color:var(--green);font-size:11px;font-weight:700;padding:4px 11px;border-radius:100px;margin-bottom:10px}
.lsb-dot{width:6px;height:6px;border-radius:50%;background:var(--green)}

/* ROLE SELECT */
.role-select{padding:4px 0}
.role-intro{font-size:13px;color:var(--text2);margin-bottom:16px}
.role-cards{display:flex;flex-direction:column;gap:12px}
.role-card{display:flex;flex-direction:column;gap:6px;background:var(--bg2);border:1.5px solid var(--border);border-radius:var(--r-lg);padding:18px;cursor:pointer;transition:all var(--ease);text-decoration:none;color:inherit}
.role-card:hover{border-color:var(--blue-border);background:var(--blue-light);transform:translateY(-2px);box-shadow:var(--sh-md)}
.rc-ico{width:40px;height:40px;border-radius:var(--r);background:#fff;border:1px solid var(--border);display:flex;align-items:center;justify-content:center;margin-bottom:4px}
.rc-title{font-size:14px;font-weight:700;color:var(--text)}
.rc-desc{font-size:12px;color:var(--text2);line-height:1.5}
.rc-arrow{font-size:12px;font-weight:600;color:var(--blue);margin-top:4px}

/* LOGIN ERROR */
.login-error{background:#fef2f2;border:1px solid #fecaca;color:#dc2626;font-size:12px;padding:10px 13px;border-radius:var(--r);margin-bottom:14px}

/* SIGNUP STEPS */
.signup-steps{display:flex;align-items:center;gap:0;margin-bottom:22px;background:var(--bg2);border:1px solid var(--border);border-radius:var(--r);padding:3px}
.ss-step{flex:1;display:flex;align-items:center;justify-content:center;gap:6px;font-size:12px;font-weight:500;color:var(--text3);padding:7px;border-radius:7px;transition:all var(--ease)}
.ss-step.active,.ss-step.current{background:#fff;color:var(--text);font-weight:600;box-shadow:var(--sh-sm)}
.ss-dot{width:19px;height:19px;border-radius:50%;background:var(--bg3);border:1.5px solid var(--border);font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;color:var(--text3);flex-shrink:0}
.ss-step.active .ss-dot{background:var(--green);border-color:var(--green);color:#fff}
.ss-step.current .ss-dot{background:var(--blue);border-color:var(--blue);color:#fff}

/* AUTH FORM */
.auth-form{display:flex;flex-direction:column}
.af-g{margin-bottom:14px}
.af-row-2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.af-g label{display:block;font-size:12px;font-weight:600;color:var(--text2);margin-bottom:5px}
.af-g input,.af-g select{width:100%;background:var(--bg2);border:1.5px solid var(--border);color:var(--text);padding:9px 12px;border-radius:var(--r);font-size:13px;transition:border-color .18s,box-shadow .18s}
.af-g input:focus,.af-g select:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.08);background:#fff}
.af-g input::placeholder{color:var(--text4)}
.af-g select option{background:#fff;color:var(--text)}
.af-lbl-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:5px}
.af-lbl-row label{margin-bottom:0}
.forgot-link{font-size:12px;color:var(--blue);font-weight:500;cursor:pointer}
.auth-switch{font-size:12px;color:var(--text3);text-align:center;margin-top:14px}
.auth-switch a{color:var(--blue);font-weight:600;cursor:pointer}
.auth-terms{font-size:11px;color:var(--text3);text-align:center;margin-top:8px;line-height:1.55}
.auth-terms a{color:var(--blue)}
.back-link{font-size:12px;color:var(--blue);cursor:pointer;text-align:center;margin-top:10px}
.auth-success{text-align:center;padding:28px 16px}
.auth-success h4{font-size:17px;font-weight:700;margin:12px 0 6px;color:var(--text)}
.auth-success p{font-size:13px;color:var(--text2)}

/* TRANSITIONS */
.modal-enter-active,.modal-leave-active{transition:all .22s ease}
.modal-enter-from,.modal-leave-to{opacity:0;transform:scale(.97)}

/* ══════════════════════════════════════════════════════════════════
   ZIRA FLOATING CHAT WIDGET
══════════════════════════════════════════════════════════════════ */
.zira-float-wrap{position:fixed;bottom:28px;right:28px;z-index:500;display:flex;flex-direction:column;align-items:flex-end;gap:12px}

/* Floating button */
.zira-float-btn{display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#1d4ed8,#2563eb);color:#fff;border:none;border-radius:100px;padding:0;cursor:pointer;box-shadow:0 8px 32px rgba(37,99,235,.45),0 2px 8px rgba(0,0,0,.1);transition:all .22s cubic-bezier(.4,0,.2,1);overflow:hidden;min-width:48px;min-height:48px}
.zira-float-btn:hover{transform:translateY(-2px);box-shadow:0 12px 40px rgba(37,99,235,.55)}
.zira-float-btn.open{width:48px;height:48px;border-radius:50%;background:linear-gradient(135deg,#374151,#1f2937)}
.zfb-content{display:flex;align-items:center;gap:10px;padding:10px 18px 10px 12px}
.zfb-avatar-wrap{position:relative;flex-shrink:0}
.zfb-av{width:32px;height:32px;border-radius:50%;background:rgba(255,255,255,.2);border:2px solid rgba(255,255,255,.3);font-size:14px;font-weight:800;display:flex;align-items:center;justify-content:center}
.zfb-pulse{position:absolute;top:-2px;right:-2px;width:10px;height:10px;border-radius:50%;background:#10b981;border:2px solid #fff;animation:zira-pulse 2s infinite}
@keyframes zira-pulse{0%,100%{box-shadow:0 0 0 0 rgba(16,185,129,.5)}50%{box-shadow:0 0 0 4px rgba(16,185,129,0)}}
.zfb-text{text-align:left}
.zfb-label{font-size:13px;font-weight:700;color:#fff;line-height:1.2}
.zfb-sub{font-size:10px;color:rgba(255,255,255,.65);font-weight:500}

/* Chat panel */
.zira-float-panel{width:340px;background:#fff;border:1px solid var(--border);border-radius:var(--r-xl);box-shadow:var(--sh-xl);display:flex;flex-direction:column;overflow:hidden;max-height:480px}
.zfp-hdr{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;background:linear-gradient(135deg,#060d28,#0f1f60);flex-shrink:0}
.zfp-hdr-left{display:flex;align-items:center;gap:10px}
.zfp-avatar{width:34px;height:34px;border-radius:10px;background:var(--blue);color:#fff;font-weight:800;font-size:15px;display:flex;align-items:center;justify-content:center;flex-shrink:0;border:2px solid rgba(255,255,255,.2)}
.zfp-name{font-size:13px;font-weight:700;color:#fff;line-height:1.2}
.zfp-status{display:flex;align-items:center;gap:5px;font-size:10px;color:rgba(255,255,255,.5)}
.zfp-online-dot{width:6px;height:6px;border-radius:50%;background:#10b981;flex-shrink:0}
.zfp-close{width:28px;height:28px;border-radius:8px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.15);color:rgba(255,255,255,.7);display:flex;align-items:center;justify-content:center;cursor:pointer;transition:all var(--ease);flex-shrink:0}
.zfp-close:hover{background:rgba(255,255,255,.2);color:#fff}

.zfp-messages{flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px;scroll-behavior:smooth;min-height:0}
.zfp-messages::-webkit-scrollbar{width:3px}
.zfp-messages::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
.zfp-msg{display:flex;align-items:flex-end;gap:7px;animation:msgIn .2s ease both}
@keyframes msgIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.zfp-msg.user{flex-direction:row-reverse}
.zfp-msg-avatar{width:24px;height:24px;border-radius:50%;background:var(--blue);color:#fff;font-size:10px;font-weight:800;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.zfp-msg-bubble{max-width:80%;padding:9px 12px;border-radius:12px;font-size:12px;line-height:1.6}
.zfp-msg.bot .zfp-msg-bubble{background:var(--bg2);border:1px solid var(--border);color:var(--text);border-radius:4px 12px 12px 12px}
.zfp-msg.user .zfp-msg-bubble{background:var(--blue);color:#fff;border-radius:12px 12px 4px 12px}

/* Typing indicator */
.zfp-typing{display:flex;align-items:center;gap:4px;padding:10px 14px}
.zfp-typing span{width:5px;height:5px;border-radius:50%;background:var(--text3);animation:typing-dot 1.2s infinite}
.zfp-typing span:nth-child(2){animation-delay:.2s}
.zfp-typing span:nth-child(3){animation-delay:.4s}
@keyframes typing-dot{0%,80%,100%{transform:scale(.8);opacity:.4}40%{transform:scale(1);opacity:1}}

/* Suggestion chips */
.zfp-suggestions{padding:0 14px 10px;display:flex;flex-wrap:wrap;gap:6px;flex-shrink:0}
.zfp-chip{background:var(--blue-light);border:1px solid var(--blue-border);color:var(--blue);font-size:11px;font-weight:600;padding:5px 10px;border-radius:100px;cursor:pointer;transition:all var(--ease);white-space:nowrap}
.zfp-chip:hover{background:var(--blue);color:#fff}

/* Input row */
.zfp-input-row{display:flex;align-items:center;gap:8px;padding:10px 12px;border-top:1px solid var(--border);flex-shrink:0;background:#fff}
.zfp-input-row input{flex:1;background:var(--bg2);border:1.5px solid var(--border);color:var(--text);padding:8px 12px;border-radius:var(--r);font-size:12px;transition:border-color .18s}
.zfp-input-row input:focus{border-color:var(--blue);background:#fff}
.zfp-input-row input::placeholder{color:var(--text3)}
.zfp-send{width:32px;height:32px;border-radius:var(--r);background:var(--blue);color:#fff;border:none;display:flex;align-items:center;justify-content:center;cursor:pointer;flex-shrink:0;transition:all var(--ease)}
.zfp-send:hover:not(:disabled){background:var(--blue2);transform:scale(1.05)}
.zfp-send:disabled{opacity:.4;cursor:default}

/* Zira panel transition */
.zira-panel-enter-active,.zira-panel-leave-active{transition:all .25s cubic-bezier(.4,0,.2,1)}
.zira-panel-enter-from,.zira-panel-leave-to{opacity:0;transform:translateY(16px) scale(.97)}

/* SCROLLBAR */
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}

/* ══════════════════════════════════════════════════════════════════
   RESPONSIVE
══════════════════════════════════════════════════════════════════ */
@media(max-width:1024px){
  .two-col{grid-template-columns:1fr;gap:44px}
  .feat-grid{grid-template-columns:1fr 1fr 1fr}
  .analytics-grid{grid-template-columns:1fr 1fr}
  .an-card.large{grid-column:span 2}
  .foot-grid{grid-template-columns:1fr;gap:36px}
  .foot-links-grid{grid-template-columns:repeat(2,1fr)}
  .cta-inner{flex-direction:column;text-align:center}
  .cta-right{align-items:center}
  .cta-note{text-align:center}
  .dropdown{min-width:340px}
}
@media(max-width:768px){
  .nav-links{display:none}
  .mob-menu{display:flex}
  .view-product-btn{display:none}
  .section{padding:60px 0}
  .container{padding:0 20px}
  .hdr{padding:0 20px}
  .hero-title{font-size:clamp(26px,7.5vw,38px)}
  .hero-sub{font-size:13px}
  .hero-btns{flex-direction:column;width:100%;max-width:270px}
  .hero-btns button{width:100%;justify-content:center}
  .hero-trust-row{flex-wrap:wrap;border-radius:var(--r-lg)}
  .ht-item{border-right:none;border-bottom:1px solid var(--border);width:50%}
  .ht-item:nth-child(odd){border-right:1px solid var(--border)}
  .ht-item:nth-last-child(-n+2){border-bottom:none}
  .dash-body{grid-template-columns:1fr}
  .dash-side{display:none}
  .dash-kpi-row{grid-template-columns:1fr 1fr}
  .dash-charts-row{grid-template-columns:1fr}
  .steps-row{grid-template-columns:1fr}
  .feat-grid{grid-template-columns:1fr 1fr}
  .analytics-grid{grid-template-columns:1fr}
  .an-card.large{grid-column:span 1}
  .cf-row-2{grid-template-columns:1fr}
  .af-row-2{grid-template-columns:1fr}
  .section-title{font-size:clamp(20px,5.5vw,28px)}
  .cta-section{padding:52px 0}
  .ms-inner{gap:8px}
  .ms-item{min-width:80px}
  .av-stats-grid{grid-template-columns:1fr 1fr}
  .mob-nav{padding:12px 20px 24px}
  .zira-float-wrap{bottom:16px;right:16px}
  .zira-float-panel{width:calc(100vw - 32px);max-width:340px}
}
@media(max-width:480px){
  .container{padding:0 16px}
  .hdr{padding:0 16px}
  .mob-nav{padding:12px 16px 24px}
  .section{padding:48px 0}
  .feat-grid{grid-template-columns:1fr}
  .foot-links-grid{grid-template-columns:repeat(2,1fr)}
  .modal-box{padding:20px 16px;border-radius:var(--r-xl) var(--r-xl) 0 0;max-width:100%;position:fixed;bottom:0;left:0;right:0;max-height:85vh}
  .modal-overlay{align-items:flex-end;padding:0}
  .av-stats-grid{grid-template-columns:1fr 1fr}
  .hero-dash{margin-top:28px}
  .ls-inner{gap:10px}
  .metrics-strip{padding:12px 0}
  .zira-float-wrap{bottom:12px;right:12px}
}
</style>
