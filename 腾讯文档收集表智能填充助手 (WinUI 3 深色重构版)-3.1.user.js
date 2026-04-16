// ==UserScript==
// @name         腾讯文档收集表智能填充助手 (WinUI 3 深色可拖拽版)
// @namespace    http://tampermonkey.net/
// @version      3.11
// @description  深色毛玻璃可拖拽面板，macOS红绿灯+强回弹动画，最大化75%视口，支持Ctrl+Enter快捷键提交，动态按钮文本
// @author       Assistant (Refactored)
// @match        *://docs.qq.com/form/page/*
// @match        *://docs.qq.com/form/fill/*
// @icon         https://docs.qq.com/favicon.ico
// @grant        GM_addStyle
// @grant        GM_setClipboard
// ==/UserScript==

(function () {
    'use strict';

    // ------------------------------ 配置常量 ------------------------------
    const CONFIG = {
        INFO_OFFICER: '范怡乐',
        CAMPUS: '麦庐',
        COLLEGE: '计算机与人工智能学院',
        STUDENT_COUNT: '45',
        DEFAULT_FEEDBACK: '老师按时到达教室，发布签到，同学们按时到达。课堂氛围活跃，无早退现象。',
        SEMESTER_START: new Date(2026, 2, 2),
        MAX_WEEK: 16,
        RANDOM_PRESETS: ['程设', '写作', '近代史', '高数', '大英', 'Java', '毛概', '形势与政策', '大物', '习概', '数电']
    };

    const EASING_BACK = 'cubic-bezier(0.68, -0.55, 0.265, 1.55)';
    const ANIM_DURATION = 600;

    // ------------------------------ 工具函数 ------------------------------
    const Utils = {
        chineseToNumber(chinese) {
            const map = { '零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };
            if (!chinese) return 0;
            if (chinese === '十') return 10;
            if (chinese.startsWith('十')) return 10 + (map[chinese[1]] || 0);
            if (chinese.endsWith('十')) return (map[chinese[0]] || 0) * 10;
            if (chinese.includes('十')) {
                const [left, right] = chinese.split('十');
                return (map[left] || 0) * 10 + (map[right] || 0);
            }
            return map[chinese] || 0;
        },

        extractOrdinal(text) {
            const match = text.match(/第([一二三四五六七八九十]+)节/);
            return match ? this.chineseToNumber(match[1]) : 1;
        },

        removeWeekInfo(text) {
            return text.replace(/第[一二三四五六七八九十]+周/g, '');
        },

        getCurrentWeek() {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const start = new Date(CONFIG.SEMESTER_START);
            start.setHours(0, 0, 0, 0);
            const diff = Math.floor((today - start) / 86400000);
            if (diff < 0) return -1;
            return Math.floor(diff / 7) + 1;
        },

        async copyToClipboard(text) {
            try {
                GM_setClipboard(text, 'text');
                return true;
            } catch {
                try {
                    await navigator.clipboard.writeText(text);
                    return true;
                } catch {
                    return false;
                }
            }
        }
    };

    // ------------------------------ 课程数据模型 ------------------------------
    class Session {
        constructor(time, classroom, order) {
            this.time = time;
            this.classroom = classroom;
            this.order = order;
        }
    }

    class Course {
        constructor(formalName, aliases, teacher, teacherCollege) {
            this.formalName = formalName;
            this.aliases = aliases;
            this.teacher = teacher;
            this.teacherCollege = teacherCollege;
            this.weeklySessions = new Map();
        }

        addSession(weekMode, time, classroom, order) {
            const weeks = [];
            if (weekMode === '全周') {
                for (let i = 1; i <= CONFIG.MAX_WEEK; i++) weeks.push(i);
            } else if (weekMode === '单周') {
                for (let i = 1; i <= CONFIG.MAX_WEEK; i += 2) weeks.push(i);
            } else if (weekMode === '双周') {
                for (let i = 2; i <= CONFIG.MAX_WEEK; i += 2) weeks.push(i);
            }
            for (const w of weeks) {
                if (!this.weeklySessions.has(w)) this.weeklySessions.set(w, []);
                const sessions = this.weeklySessions.get(w);
                if (!sessions.some(s => s.time === time && s.classroom === classroom)) {
                    sessions.push(new Session(time, classroom, order));
                }
            }
            for (const [_, list] of this.weeklySessions) {
                list.sort((a, b) => a.order - b.order);
            }
        }
    }

    // ------------------------------ 课程仓库 ------------------------------
    class CourseRepository {
        constructor() {
            this.courses = new Map();
            this._buildData();
        }

        _buildData() {
            const add = (name, aliases, teacher, college) => {
                const c = new Course(name, aliases, teacher, college);
                this.courses.set(name, c);
                return c;
            };

            const c1 = add('程序设计实践', ['easyx', '程设', '程设实践', '程序设计实践'], '焦贤沛', '计算机与人工智能学院');
            const c2 = add('写作与沟通I', ['语文', '写作', '沟通', '写沟', '写作与沟通'], '王柳芳', '社会与人文学院');
            const c3 = add('中国近现代史纲要', ['近代史', '近现代史', '史纲', '纲要'], '吴通福', '马克思主义学院');
            const c4 = add('高等数学II', ['高数II', '高数2', '高等数学二', '高数二', '高数', '高等数学'], '俞丽兰', '信息管理与数学学院');
            const c5 = add('大学英语II', ['大英', '英语', '大英II', '英语II', '大学英语'], '史希平', '外国语学院');
            const c6 = add('体育2', ['体育', '体育二'], '彭永善', '体育学院');
            const c7 = add('面向对象程序设计(双语)', ['java', 'oop', '面向对象', '面向对象程序设计'], '夏雪', '计算机与人工智能学院');
            const c8 = add('毛泽东思想和中国特色社会主义理论体系概论', ['毛概', '毛中特', '毛泽东', '理论体系'], '康立芳', '马克思主义学院');
            const c9 = add('形势与政策II', ['形策', '形势与政策'], '谢尔艾力.库尔班', '马克思主义学院');
            const c10 = add('大学物理', ['大物', '物理'], '余泉茂', '软件与物联网工程学院');
            const c11 = add('习近平新时代中国特色社会主义思想概论', ['习概', '新思想', '习近平'], '徐腊梅', '马克思主义学院');
            const c12 = add('数字逻辑与数字系统', ['数逻', '数字逻辑', '数电'], '包晗秋', '计算机与人工智能学院');

            c1.addSession('全周', '3-12', '图文楼M103', 1);
            c2.addSession('全周', '1-34', '3310', 1);
            c3.addSession('全周', '2-345', '3203', 1);
            c4.addSession('全周', '1-101112', '3303', 1);
            c4.addSession('全周', '3-345', '3304', 2);
            c5.addSession('全周', '4-34', '3207', 1);
            c6.addSession('全周', '5-34', '乒乓球场T021', 1);
            c7.addSession('单周', '1-678', '3407', 1);
            c7.addSession('双周', '1-678', '图文楼M106', 1);
            c8.addSession('全周', '2-678', '3303', 1);
            c9.addSession('全周', '3-67', '3403', 1);
            c10.addSession('全周', '4-678', '二教2103', 1);
            c11.addSession('全周', '5-678', '3202', 1);
            c12.addSession('全周', '2-1011', '3313', 1);
            c12.addSession('双周', '4-1011', '图文楼M102', 2);
        }

        findCourse(rawText) {
            const lower = rawText.toLowerCase();
            for (const course of this.courses.values()) {
                if (course.aliases.some(alias => lower.includes(alias.toLowerCase()))) {
                    return course;
                }
            }
            return null;
        }
    }

    // ------------------------------ 解析服务 ------------------------------
    class CourseParser {
        constructor(repo) {
            this.repo = repo;
        }

        parse(inputText, customFeedback = null) {
            const cleaned = Utils.removeWeekInfo(inputText);
            const ordinal = Utils.extractOrdinal(cleaned);
            const course = this.repo.findCourse(cleaned);
            if (!course) return { error: '未匹配到课程，请检查课程别名' };

            const week = Utils.getCurrentWeek();
            if (week === -1) return { error: '当前日期早于开学日' };
            if (week > CONFIG.MAX_WEEK) return { error: `当前第${week}周超出课表范围` };

            const sessions = course.weeklySessions.get(week);
            if (!sessions || sessions.length === 0) {
                return { error: `「${course.formalName}」第${week}周无课` };
            }
            if (ordinal < 1 || ordinal > sessions.length) {
                return { error: `第${week}周共${sessions.length}节，您指定第${ordinal}节无效` };
            }

            const selected = sessions[ordinal - 1];
            const feedback = customFeedback?.trim() || CONFIG.DEFAULT_FEEDBACK;

            return {
                college: CONFIG.COLLEGE,
                officer: CONFIG.INFO_OFFICER,
                campus: CONFIG.CAMPUS,
                classroom: selected.classroom,
                time: selected.time,
                teacher: course.teacher,
                teacherCollege: course.teacherCollege,
                courseName: course.formalName,
                studentCount: CONFIG.STUDENT_COUNT,
                feedback,
                week
            };
        }
    }

    // ------------------------------ 表单填充器 ------------------------------
    class FormFiller {
        fillField(questionTitle, value) {
            if (!value) return false;
            const questions = document.querySelectorAll('.question');
            for (const q of questions) {
                let titleSpan = q.querySelector('.question-title .form-auto-ellipsis');
                let titleText = titleSpan ? titleSpan.textContent.trim() : '';
                if (!titleText) {
                    const titleDiv = q.querySelector('.question-title');
                    if (titleDiv) titleText = titleDiv.textContent.trim().replace(/\*$/, '').trim();
                }
                if (titleText === questionTitle) {
                    const input = q.querySelector('textarea') || q.querySelector('input[type="text"], input:not([type])');
                    if (input) {
                        input.focus();
                        input.value = value;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                        input.dispatchEvent(new Event('blur', { bubbles: true }));
                        input.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true }));
                        input.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
                        return true;
                    }
                    return false;
                }
            }
            return false;
        }

        fillAll(fieldMap) {
            let success = true;
            for (const [title, val] of fieldMap) {
                if (!this.fillField(title, val)) success = false;
            }
            return success;
        }
    }

    // ------------------------------ 深色毛玻璃可拖拽面板 ------------------------------
    class DraggableWinUIPanel {
        constructor(onFill) {
            this.onFill = onFill;
            this.panel = null;
            this.btn = null;
            this.originalBtnText = '✨ 填写并复制';
            this.randomBtnText = '🎲 随机选择课程';
            this.timeoutId = null;
            this.isDragging = false;
            this.startX = 0;
            this.startY = 0;
            this.initialLeft = 0;
            this.initialTop = 0;
            this.isMaximized = false;
            this.normalStyle = { width: 340, height: null, left: null, top: null, right: null, bottom: null };
            this.miniButton = null;
            this.isAnimating = false;
        }

        create() {
            this.panel = document.createElement('div');
            this.panel.id = 'winui-draggable-panel';
            this.panel.innerHTML = `
                <div class="winui-panel">
                    <div class="winui-title-bar">
                        <div class="window-controls">
                            <div class="control close" id="winui-close">${this._closeIconSVG()}</div>
                            <div class="control minimize" id="winui-minimize">${this._minimizeIconSVG()}</div>
                            <div class="control maximize" id="winui-maximize">${this._maximizeIconSVG()}</div>
                        </div>
                        <div class="winui-title" id="drag-handle">📋 课程信息填充</div>
                        <div class="title-spacer"></div>
                    </div>
                    <div class="winui-content">
                        <div class="winui-input-field">
                            <label>课程描述 <span class="optional">含节次</span></label>
                            <input type="text" id="courseDescInput" placeholder="例：第二节的高数II（留空则随机）" value="">
                        </div>
                        <div class="winui-input-field">
                            <label>反馈内容</label>
                            <textarea id="feedbackInput" rows="3">${CONFIG.DEFAULT_FEEDBACK}</textarea>
                        </div>
                        <button id="fillBtn" class="winui-button accent">${this.originalBtnText}</button>
                    </div>
                </div>
            `;
            document.body.appendChild(this.panel);
            this.btn = document.getElementById('fillBtn');
            this._injectStyles();
            this._restorePosition();
            this._setupDragging();
            this._setupWindowControls();
            this._bindEvents();
            this._createMiniButton();

            requestAnimationFrame(() => {
                const rect = this.panel.getBoundingClientRect();
                this.normalStyle.height = rect.height;
            });
        }

        _updateButtonText() {
            const descInput = document.getElementById('courseDescInput');
            if (descInput && descInput.value.trim() === '') {
                this.btn.innerHTML = this.randomBtnText;
            } else {
                this.btn.innerHTML = this.originalBtnText;
            }
        }

        _closeIconSVG() {
            return `<svg width="15" height="15" viewBox="0 0 15 15" fill="none"><circle cx="6" cy="6" r="6" fill="#FF5F57" stroke="#E0443E" stroke-width="0.5"/></svg>`;
        }

        _minimizeIconSVG() {
            return `<svg width="15" height="15" viewBox="0 0 15 15" fill="none"><circle cx="6" cy="6" r="6" fill="#FFBD2E" stroke="#DFA123" stroke-width="0.5"/></svg>`;
        }

        _maximizeIconSVG() {
            return `<svg width="15" height="15" viewBox="0 0 15 15" fill="none"><circle cx="6" cy="6" r="6" fill="#28C840" stroke="#1BAC2C" stroke-width="0.5"/></svg>`;
        }

        _injectStyles() {
            GM_addStyle(`
                :root {
                    --panel-bg: rgba(255, 255, 255, 0.5);
                    --panel-border: rgba(255, 255, 255, 0.2);
                    --text-primary: #1f1f1f;
                    --text-secondary: #5c5c5c;
                    --input-bg: rgba(255, 255, 255, 0.4);
                    --input-border: rgba(0, 0, 0, 0.1);
                    --input-focus-border: #0078d4;
                    --input-focus-shadow: rgba(0, 120, 212, 0.3);
                    --button-bg: rgba(0, 0, 0, 0.04);
                    --button-border: rgba(0, 0, 0, 0.1);
                    --button-accent-bg: #0078d4;
                    --button-accent-hover: #106ebe;
                    --button-accent-active: #005a9e;
                    --toast-bg-error: rgba(196, 43, 28, 0.9);
                }
                @media (prefers-color-scheme: dark) {
                    :root {
                        --panel-bg: rgba(32, 32, 32, 0.5);
                        --panel-border: rgba(255, 255, 255, 0.08);
                        --text-primary: #ffffff;
                        --text-secondary: #c7c7c7;
                        --input-bg: rgba(50, 50, 50, 0.5);
                        --input-border: rgba(255, 255, 255, 0.1);
                        --input-focus-border: #60cdff;
                        --input-focus-shadow: rgba(96, 205, 255, 0.3);
                        --button-bg: rgba(255, 255, 255, 0.06);
                        --button-border: rgba(255, 255, 255, 0.1);
                        --button-accent-bg: #60cdff;
                        --button-accent-hover: #3aa0d0;
                        --button-accent-active: #1e7aa8;
                        --toast-bg-error: rgba(220, 60, 60, 0.9);
                    }
                    .winui-button.accent { color: #1a1a1a !important; }
                }

                #winui-draggable-panel {
                    position: fixed;
                    width: 340px;
                    height: auto;
                    z-index: 10000;
                    backdrop-filter: blur(20px) saturate(180%);
                    -webkit-backdrop-filter: blur(20px) saturate(180%);
                    font-family: 'Segoe UI Variable', 'Segoe UI', system-ui, sans-serif;
                    pointer-events: none;
                    will-change: transform, opacity, width, left, top;
                    border-radius: 12px;
                }
                .winui-panel {
                    background: var(--panel-bg);
                    border-radius: 12px;
                    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2), 0 0 0 1px var(--panel-border) inset;
                    border: 1px solid rgba(255, 255, 255, 0.05);
                    overflow: hidden;
                    pointer-events: auto;
                    color: var(--text-primary);
                    display: flex;
                    flex-direction: column;
                    height: 100%;
                    min-height: 0;
                }
                .winui-title-bar {
                    display: flex;
                    align-items: center;
                    padding: 13px;
                    flex-shrink: 0;
                }
                .window-controls {
                    display: flex;
                    gap: 4px;
                }
                .control {
                    width: 16px;
                    height: 16px;
                    border-radius: 50%;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    cursor: pointer;
                    transition: filter 0.2s;
                }
                .control:hover { filter: brightness(0.9); }
                .winui-title {
                    flex: 1;
                    font-size: 14px;
                    font-weight: 600;
                    color: var(--text-primary);
                    letter-spacing: -0.01em;
                    cursor: move;
                    user-select: none;
                    text-align: center;
                }
                .title-spacer { width: 48px; }
                .winui-content {
                    padding: 0 20px 20px;
                    flex: 1;
                    display: flex;
                    flex-direction: column;
                    min-height: 0;
                }
                .winui-input-field {
                    margin-bottom: 16px;
                }
                .winui-input-field label {
                    display: block;
                    font-size: 13px;
                    font-weight: 500;
                    margin-bottom: 6px;
                    color: var(--text-primary);
                }
                .winui-input-field .optional {
                    font-weight: 400;
                    color: var(--text-secondary);
                    margin-left: 6px;
                }
                #feedbackInput {
                    margin-bottom: 16px;
                }
                .winui-input-field input,
                .winui-input-field textarea {
                    width: 100%;
                    padding: 10px 12px;
                    background: var(--input-bg);
                    border: 1px solid var(--input-border);
                    border-radius: 8px;
                    font-size: 14px;
                    color: var(--text-primary);
                    outline: none;
                    box-sizing: border-box;
                    transition: border 0.2s, box-shadow 0.2s, background 0.2s;
                    font-family: inherit;
                    resize: none;
                    overflow: auto;
                }
                .winui-input-field input:focus,
                .winui-input-field textarea:focus {
                    border-color: var(--input-focus-border);
                    box-shadow: 0 0 0 3px var(--input-focus-shadow);
                    background: var(--input-bg);
                }
                .winui-input-field:last-of-type {
                    flex: 1;
                    display: flex;
                    flex-direction: column;
                    margin-bottom: 0;
                }
                .winui-input-field:last-of-type textarea {
                    flex: 1;
                    resize: none;
                    min-height: 80px;
                }
                .winui-button {
                    width: 100%;
                    padding: 10px 16px;
                    border-radius: 8px;
                    border: none;
                    font-weight: 500;
                    font-size: 14px;
                    cursor: pointer;
                    transition: all 0.25s ${EASING_BACK};
                    background: var(--button-bg);
                    color: var(--text-primary);
                    border: 1px solid var(--button-border);
                }
                .winui-button.accent {
                    background: var(--button-accent-bg);
                    border: none;
                    margin-top: auto;
                    flex-shrink: 0;
                }
                .winui-button.accent:hover { background: var(--button-accent-hover); }
                .winui-button.accent:active { background: var(--button-accent-active); transform: scale(0.97); }
                .winui-button.accent.success {
                    background: #0e7a4d !important;
                    color: white !important;
                }
                #winui-mini-button {
                    position: fixed;
                    bottom: 30px;
                    right: 30px;
                    width: 48px;
                    height: 48px;
                    border-radius: 24px;
                    background: var(--panel-bg);
                    backdrop-filter: blur(20px);
                    -webkit-backdrop-filter: blur(20px);
                    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
                    border: 1px solid var(--panel-border);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    cursor: pointer;
                    z-index: 10001;
                    transition: opacity 0.2s, transform 0.2s ${EASING_BACK};
                    opacity: 0;
                    pointer-events: none;
                }
                #winui-mini-button.visible {
                    opacity: 1;
                    pointer-events: auto;
                }
                #winui-mini-button:hover { transform: scale(1.05); }
                #winui-mini-button svg { width: 24px; height: 24px; fill: var(--text-primary); }
                @media (max-width: 700px) {
                    #winui-draggable-panel {
                        width: calc(100% - 20px) !important;
                        left: 10px !important;
                        right: 10px !important;
                        top: auto !important;
                        bottom: 10px !important;
                    }
                    .winui-title { cursor: default; }
                }
            `);
        }

        _restorePosition() {
            const savedLeft = localStorage.getItem('winuiPanelLeft');
            const savedTop = localStorage.getItem('winuiPanelTop');
            if (savedLeft && savedTop) {
                this.panel.style.left = savedLeft + 'px';
                this.panel.style.top = savedTop + 'px';
                this.panel.style.right = 'auto';
                this.panel.style.bottom = 'auto';
            } else {
                this.panel.style.top = '100px';
                this.panel.style.right = '20px';
                this.panel.style.left = 'auto';
                this.panel.style.bottom = 'auto';
            }
            this.panel.offsetHeight;
            const rect = this.panel.getBoundingClientRect();
            this.normalStyle.left = rect.left;
            this.normalStyle.top = rect.top;
            this.normalStyle.width = rect.width;
            this.normalStyle.height = rect.height;
        }

        _setupDragging() {
            const header = document.getElementById('drag-handle');
            if (!header) return;

            header.addEventListener('mousedown', (e) => {
                if (e.button !== 0) return;
                if (window.innerWidth <= 700 || this.isAnimating) return;
                this.isDragging = true;
                this.startX = e.clientX;
                this.startY = e.clientY;
                const rect = this.panel.getBoundingClientRect();
                this.initialLeft = rect.left;
                this.initialTop = rect.top;
                this.panel.style.transition = 'none';
                this.panel.style.right = 'auto';
                this.panel.style.left = this.initialLeft + 'px';
                this.panel.style.top = this.initialTop + 'px';
                e.preventDefault();
            });

            const onMouseMove = (e) => {
                if (!this.isDragging) return;
                const dx = e.clientX - this.startX;
                const dy = e.clientY - this.startY;
                let newLeft = this.initialLeft + dx;
                let newTop = this.initialTop + dy;
                const maxX = window.innerWidth - this.panel.offsetWidth;
                const maxY = window.innerHeight - this.panel.offsetHeight;
                newLeft = Math.min(Math.max(0, newLeft), maxX);
                newTop = Math.min(Math.max(0, newTop), maxY);
                this.panel.style.left = newLeft + 'px';
                this.panel.style.top = newTop + 'px';
            };

            const onMouseUp = () => {
                if (this.isDragging) {
                    this.isDragging = false;
                    this.panel.style.transition = '';
                    const rect = this.panel.getBoundingClientRect();
                    localStorage.setItem('winuiPanelLeft', rect.left);
                    localStorage.setItem('winuiPanelTop', rect.top);
                    if (!this.isMaximized) {
                        this.normalStyle.left = rect.left;
                        this.normalStyle.top = rect.top;
                        this.normalStyle.width = rect.width;
                        this.normalStyle.height = rect.height;
                    }
                }
            };

            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        }

        _setupWindowControls() {
            document.getElementById('winui-close').addEventListener('click', () => this._animateClose());
            document.getElementById('winui-minimize').addEventListener('click', () => this._animateMinimize());
            document.getElementById('winui-maximize').addEventListener('click', () => this._toggleMaximize());
        }

        _animateClose() {
            if (this.isAnimating) return;
            this.isAnimating = true;
            const panel = this.panel;

            panel.style.visibility = 'visible';
            panel.style.transition = '';
            panel.style.transform = '';
            panel.style.opacity = '';
            panel.offsetHeight;

            panel.style.pointerEvents = 'none';
            panel.style.transformOrigin = 'center center';
            panel.style.transition = `transform ${ANIM_DURATION}ms ${EASING_BACK}, opacity ${ANIM_DURATION}ms`;
            panel.style.transform = 'scale(0)';
            panel.style.opacity = '0';
            setTimeout(() => {
                panel.remove();
                if (this.miniButton) this.miniButton.remove();
                this.isAnimating = false;
            }, ANIM_DURATION);
        }

        _animateMinimize() {
            if (this.isAnimating) return;
            this.isAnimating = true;
            const panelRect = this.panel.getBoundingClientRect();
            const miniRect = this.miniButton.getBoundingClientRect();

            const scaleX = miniRect.width / panelRect.width;
            const scaleY = miniRect.height / panelRect.height;
            const translateX = miniRect.left + miniRect.width / 2 - (panelRect.left + panelRect.width / 2);
            const translateY = miniRect.top + miniRect.height / 2 - (panelRect.top + panelRect.height / 2);

            this.panel.style.transition = `transform ${ANIM_DURATION}ms ${EASING_BACK}, opacity ${ANIM_DURATION}ms`;
            this.panel.style.transform = `translate(${translateX}px, ${translateY}px) scale(${scaleX}, ${scaleY})`;
            this.panel.style.opacity = '0.6';

            setTimeout(() => {
                this.panel.style.visibility = 'hidden';
                this.panel.style.transition = '';
                this.panel.style.transform = '';
                this.panel.style.opacity = '';
                this.miniButton.classList.add('visible');
                this.isAnimating = false;
            }, ANIM_DURATION);
        }

        _animateRestoreFromMini() {
            if (this.isAnimating) return;
            this.isAnimating = true;

            this.panel.style.visibility = 'visible';
            const panelRect = this.panel.getBoundingClientRect();
            const miniRect = this.miniButton.getBoundingClientRect();

            const scaleX = miniRect.width / panelRect.width;
            const scaleY = miniRect.height / panelRect.height;
            const translateX = miniRect.left + miniRect.width / 2 - (panelRect.left + panelRect.width / 2);
            const translateY = miniRect.top + miniRect.height / 2 - (panelRect.top + panelRect.height / 2);

            this.panel.style.transition = 'none';
            this.panel.style.transform = `translate(${translateX}px, ${translateY}px) scale(${scaleX}, ${scaleY})`;
            this.panel.style.opacity = '0.6';
            this.panel.offsetHeight;

            this.panel.style.transition = `transform ${ANIM_DURATION}ms ${EASING_BACK}, opacity ${ANIM_DURATION}ms`;
            this.panel.style.transform = 'translate(0, 0) scale(1)';
            this.panel.style.opacity = '1';
            this.miniButton.classList.remove('visible');

            setTimeout(() => {
                this.panel.style.transition = '';
                this.panel.style.transform = '';
                this.isAnimating = false;
            }, ANIM_DURATION);
        }

        _toggleMaximize() {
            if (this.isAnimating) return;
            const rect = this.panel.getBoundingClientRect();

            const computedHeight = getComputedStyle(this.panel).height;
            if (!this.panel.style.height || computedHeight === 'auto') {
                this.panel.style.height = rect.height + 'px';
                this.panel.offsetHeight;
            }

            if (!this.isMaximized) {
                this.normalStyle.width = rect.width;
                this.normalStyle.height = rect.height;
                this.normalStyle.left = rect.left;
                this.normalStyle.top = rect.top;
                this.normalStyle.right = this.panel.style.right;

                const newWidth = window.innerWidth * 0.75;
                const newHeight = window.innerHeight * 0.75;
                const newLeft = (window.innerWidth - newWidth) / 2;
                const newTop = (window.innerHeight - newHeight) / 2;

                this.panel.style.transition = `width ${ANIM_DURATION}ms ${EASING_BACK}, height ${ANIM_DURATION}ms ${EASING_BACK}, left ${ANIM_DURATION}ms ${EASING_BACK}, top ${ANIM_DURATION}ms ${EASING_BACK}`;
                this.panel.style.width = newWidth + 'px';
                this.panel.style.height = newHeight + 'px';
                this.panel.style.left = newLeft + 'px';
                this.panel.style.top = newTop + 'px';
                this.panel.style.right = 'auto';
                this.isMaximized = true;
            } else {
                this.panel.style.transition = `width ${ANIM_DURATION}ms ${EASING_BACK}, height ${ANIM_DURATION}ms ${EASING_BACK}, left ${ANIM_DURATION}ms ${EASING_BACK}, top ${ANIM_DURATION}ms ${EASING_BACK}`;
                this.panel.style.width = this.normalStyle.width + 'px';
                this.panel.style.height = this.normalStyle.height + 'px';
                this.panel.style.left = this.normalStyle.left + 'px';
                this.panel.style.top = this.normalStyle.top + 'px';
                this.panel.style.right = 'auto';
                this.isMaximized = false;
            }

            setTimeout(() => {
                this.panel.style.transition = '';
                if (!this.isMaximized) {
                    const newRect = this.panel.getBoundingClientRect();
                    this.normalStyle.width = newRect.width;
                    this.normalStyle.height = newRect.height;
                }
            }, ANIM_DURATION);
        }

        _createMiniButton() {
            this.miniButton = document.createElement('div');
            this.miniButton.id = 'winui-mini-button';
            this.miniButton.innerHTML = `<svg viewBox="0 0 24 24"><path d="M20 2H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h14l4 4V4c0-1.1-.9-2-2-2z"/></svg>`;
            document.body.appendChild(this.miniButton);

            this.miniButton.addEventListener('click', () => {
                if (this.isAnimating) return;
                this._animateRestoreFromMini();
            });

            let isDraggingMini = false, startX, startY, startLeft, startTop;
            this.miniButton.addEventListener('mousedown', (e) => {
                if (e.button !== 0) return;
                isDraggingMini = true;
                startX = e.clientX;
                startY = e.clientY;
                const rect = this.miniButton.getBoundingClientRect();
                startLeft = rect.left;
                startTop = rect.top;
                this.miniButton.style.transition = 'none';
                e.preventDefault();
            });
            document.addEventListener('mousemove', (e) => {
                if (!isDraggingMini) return;
                const dx = e.clientX - startX;
                const dy = e.clientY - startY;
                let newLeft = startLeft + dx;
                let newTop = startTop + dy;
                newLeft = Math.min(window.innerWidth - 48, Math.max(0, newLeft));
                newTop = Math.min(window.innerHeight - 48, Math.max(0, newTop));
                this.miniButton.style.left = newLeft + 'px';
                this.miniButton.style.top = newTop + 'px';
                this.miniButton.style.right = 'auto';
                this.miniButton.style.bottom = 'auto';
            });
            document.addEventListener('mouseup', () => {
                if (isDraggingMini) {
                    isDraggingMini = false;
                    this.miniButton.style.transition = '';
                }
            });
        }

        _bindEvents() {
            const descInput = document.getElementById('courseDescInput');
            const fbInput = document.getElementById('feedbackInput');

            // 实时更新按钮文本
            const updateBtnText = () => this._updateButtonText();
            descInput.addEventListener('input', updateBtnText);
            // 初始调用一次
            updateBtnText();

            // Ctrl+Enter 快捷键
            const handleCtrlEnter = (e) => {
                if (e.ctrlKey && e.key === 'Enter') {
                    e.preventDefault();
                    this.btn.click();
                }
            };
            descInput.addEventListener('keydown', handleCtrlEnter);
            fbInput.addEventListener('keydown', handleCtrlEnter);

            this.btn.addEventListener('click', async () => {
                const desc = descInput.value.trim();
                if (!desc) {
                    descInput.value = CONFIG.RANDOM_PRESETS[Math.floor(Math.random() * CONFIG.RANDOM_PRESETS.length)];

                    // 手动触发一次 input 事件以更新按钮文本（实际上输入值变化会自动触发，但为了保险）
                    descInput.dispatchEvent(new Event('input', { bubbles: true }));
                    return;
                }

                const fb = fbInput.value;
                this.btn.disabled = true;
                try {
                    await this.onFill(desc, fb);
                } finally {
                    this.btn.disabled = false;
                }
            });
        }

        showSuccessOnButton() {
            if (this.timeoutId) clearTimeout(this.timeoutId);
            this.btn.classList.add('success');
            this.btn.innerHTML = `<span style="display:flex;align-items:center;justify-content:center;gap:6px;">✅ 已复制到剪贴板</span>`;
            this.timeoutId = setTimeout(() => {
                this.btn.classList.remove('success');
                this._updateButtonText(); // 恢复原始按钮文本
                this.timeoutId = null;
            }, 1800);
        }

        showErrorToast(message) {
            const toast = document.createElement('div');
            toast.textContent = message;
            toast.style.cssText = `
                position: fixed; bottom: 30px; right: 30px;
                background: var(--toast-bg-error); color: white;
                padding: 12px 20px; border-radius: 8px; font-size: 14px;
                z-index: 100000; box-shadow: 0 8px 20px rgba(0,0,0,0.3);
                font-family: 'Segoe UI Variable', system-ui, sans-serif;
                backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
                border: 1px solid rgba(255,255,255,0.2);
                transition: opacity 0.3s; pointer-events: none;
            `;
            document.body.appendChild(toast);
            setTimeout(() => {
                toast.style.opacity = '0';
                setTimeout(() => toast.remove(), 500);
            }, 2500);
        }
    }

    // ------------------------------ 主应用 ------------------------------
    class App {
        constructor() {
            this.repo = new CourseRepository();
            this.parser = new CourseParser(this.repo);
            this.filler = new FormFiller();
            this.panel = new DraggableWinUIPanel(this._handleFill.bind(this));
            this.init();
        }

        async _handleFill(courseDesc, feedbackText) {
            const data = this.parser.parse(courseDesc, feedbackText);
            if (data.error) {
                this.panel.showErrorToast(data.error);
                return;
            }

            const fields = [
                ['教师所在学院', data.teacherCollege],
                ['教师', data.teacher],
                ['课程', data.courseName],
                ['校区', data.campus],
                ['上课时间', data.time],
                ['课表教室', data.classroom],
                ['应到人数', data.studentCount],
                ['实到人数', data.studentCount],
                ['反馈内容', data.feedback],
                ['周次', data.week.toString()],
                ['信息员', data.officer],
                ['信息员所在学院', data.college]
            ];

            this.filler.fillAll(fields);

            const fullText = `学院：${data.college}
信息员：${data.officer}
校区：${data.campus}
课表教室 ：${data.classroom}
时间：${data.time}
教师：${data.teacher}
教师所在学院：${data.teacherCollege}
课程：${data.courseName}
人数：(${data.studentCount}/${data.studentCount})
反馈内容：${data.feedback}`;

            const copied = await Utils.copyToClipboard(fullText);
            if (copied) {
                this.panel.showSuccessOnButton();
            } else {
                this.panel.showErrorToast('复制失败，请手动复制');
            }
        }

        init() {
            if (document.querySelector('.question')) {
                this.panel.create();
            } else {
                const observer = new MutationObserver((_, obs) => {
                    if (document.querySelector('.question')) {
                        obs.disconnect();
                        this.panel.create();
                    }
                });
                observer.observe(document.body, { childList: true, subtree: true });
                setTimeout(() => {
                    observer.disconnect();
                    if (!document.querySelector('.question')) {
                        this.panel.create();
                    }
                }, 10000);
            }
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => new App());
    } else {
        new App();
    }
})();