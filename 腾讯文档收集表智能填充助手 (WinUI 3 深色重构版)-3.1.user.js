// ==UserScript==
// @name         腾讯文档收集表智能填充助手 (WinUI 3 深色可拖拽版)
// @namespace    http://tampermonkey.net/
// @version      3.2
// @description  深色毛玻璃可拖拽面板，复制成功按钮动画反馈
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
        INFO_OFFICER: "范怡乐",
        CAMPUS: "麦庐",
        COLLEGE: "计算机与人工智能学院",
        STUDENT_COUNT: "45",
        DEFAULT_FEEDBACK: "老师按时到达教室，发布签到，同学们按时到达。课堂氛围活跃，无早退现象。",
        SEMESTER_START: new Date(2026, 2, 2),
        MAX_WEEK: 16
    };

    // ------------------------------ 工具函数 ------------------------------
    class Utils {
        static chineseToNumber(chinese) {
            const map = { '零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };
            if (!chinese) return 0;
            if (chinese === '十') return 10;
            if (chinese.startsWith('十')) return 10 + (map[chinese[1]] || 0);
            if (chinese.endsWith('十')) return (map[chinese[0]] || 0) * 10;
            if (chinese.includes('十')) {
                let [left, right] = chinese.split('十');
                return (map[left] || 0) * 10 + (map[right] || 0);
            }
            return map[chinese] || 0;
        }

        static extractOrdinal(text) {
            const match = text.match(/第([一二三四五六七八九十]+)节/);
            return match ? Utils.chineseToNumber(match[1]) : 1;
        }

        static removeWeekInfo(text) {
            return text.replace(/第[一二三四五六七八九十]+周/g, '');
        }

        static getCurrentWeek() {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const start = new Date(CONFIG.SEMESTER_START);
            start.setHours(0, 0, 0, 0);
            const diff = Math.floor((today - start) / 86400000);
            if (diff < 0) return -1;
            return Math.floor(diff / 7) + 1;
        }

        static async copyToClipboard(text) {
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
    }

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
            if (weekMode === "全周") {
                for (let i = 1; i <= CONFIG.MAX_WEEK; i++) weeks.push(i);
            } else if (weekMode === "单周") {
                for (let i = 1; i <= CONFIG.MAX_WEEK; i += 2) weeks.push(i);
            } else if (weekMode === "双周") {
                for (let i = 2; i <= CONFIG.MAX_WEEK; i += 2) weeks.push(i);
            }
            for (let w of weeks) {
                if (!this.weeklySessions.has(w)) this.weeklySessions.set(w, []);
                const sessions = this.weeklySessions.get(w);
                if (!sessions.some(s => s.time === time && s.classroom === classroom)) {
                    sessions.push(new Session(time, classroom, order));
                }
            }
            for (let [_, list] of this.weeklySessions) {
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

            const c1 = add("程序设计实践", ["Easyx", "easyx", "EasyX", "程设", "程设实践", "程序设计实践"], "焦贤沛", "计算机与人工智能学院");
            const c2 = add("写作与沟通I", ["语文", "写作", "沟通", "写沟", "写作与沟通"], "王柳芳", "社会与人文学院");
            const c3 = add("中国近现代史纲要", ["近代史", "近现代史", "史纲", "纲要"], "吴通福", "马克思主义学院");
            const c4 = add("高等数学II", ["高数II", "高数2", "高等数学二", "高数二", "高数", "高等数学"], "俞丽兰", "信息管理与数学学院");
            const c5 = add("大学英语II", ["大英", "英语", "大英II", "英语II", "大学英语"], "史希平", "外国语学院");
            const c6 = add("体育2", ["体育", "体育二"], "彭永善", "体育学院");
            const c7 = add("面向对象程序设计(双语)", ["Java", "JAVA", "java", "OOP", "面向对象", "面向对象程序设计"], "夏雪", "计算机与人工智能学院");
            const c8 = add("毛泽东思想和中国特色社会主义理论体系概论", ["毛概", "毛中特", "毛泽东", "理论体系"], "康立芳", "马克思主义学院");
            const c9 = add("形势与政策II", ["形策", "形势与政策"], "谢尔艾力.库尔班", "马克思主义学院");
            const c10 = add("大学物理", ["大物", "物理"], "余泉茂", "软件与物联网工程学院");
            const c11 = add("习近平新时代中国特色社会主义思想概论", ["习概", "新思想概论", "习近平"], "徐腊梅", "马克思主义学院");
            const c12 = add("数字逻辑与数字系统", ["数逻", "数字逻辑", "数电"], "包晗秋", "计算机与人工智能学院");

            c1.addSession("全周", "3-12", "图文楼M103", 1);
            c2.addSession("全周", "1-34", "3310", 1);
            c3.addSession("全周", "2-345", "3203", 1);
            c4.addSession("全周", "1-101112", "3303", 1);
            c4.addSession("全周", "3-345", "3304", 2);
            c5.addSession("全周", "4-34", "3207", 1);
            c6.addSession("全周", "5-34", "乒乓球场T021", 1);
            c7.addSession("单周", "1-678", "3407", 1);
            c7.addSession("双周", "1-678", "图文楼M106", 1);
            c8.addSession("全周", "2-678", "3303", 1);
            c9.addSession("全周", "3-67", "3403", 1);
            c10.addSession("全周", "4-678", "二教2103", 1);
            c11.addSession("全周", "5-678", "3202", 1);
            c12.addSession("全周", "2-1011", "3313", 1);
            c12.addSession("双周", "4-1011", "图文楼M102", 2);
        }

        findCourse(rawText) {
            const lower = rawText.toLowerCase();
            for (let course of this.courses.values()) {
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
            if (!course) return { error: `未匹配到课程，请检查课程别名` };

            const week = Utils.getCurrentWeek();
            if (week === -1) return { error: `当前日期早于开学日` };
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
                feedback: feedback,
                week: week
            };
        }
    }

    // ------------------------------ 表单填充器 ------------------------------
    class FormFiller {
        fillField(questionTitle, value) {
            if (!value) return false;
            const questions = document.querySelectorAll('.question');
            for (let q of questions) {
                const titleSpan = q.querySelector('.question-title .form-auto-ellipsis');
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
            for (let [title, val] of fieldMap) {
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
            this.timeoutId = null;
            this.isDragging = false;
            this.startX = 0;
            this.startY = 0;
            this.initialLeft = 0;
            this.initialTop = 0;
        }

        create() {
            this.panel = document.createElement('div');
            this.panel.id = 'winui-draggable-panel';
            this.panel.innerHTML = `
                <div class="winui-panel">
                    <div class="winui-title" id="drag-handle">📋 课程信息填充</div>
                    <div class="winui-content">
                        <div class="winui-input-field">
                            <label>课程描述 <span class="optional">含节次</span></label>
                            <input type="text" id="courseDescInput" placeholder="例：第二节的高数II" value="第二节的高数II">
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
            this._bindEvents();
        }

        _injectStyles() {
            GM_addStyle(`
                /* 基础变量 - 浅色模式 */
                :root {
                    --panel-bg: rgba(255, 255, 255, 0.75);
                    --panel-border: rgba(255, 255, 255, 0.3);
                    --text-primary: #1f1f1f;
                    --text-secondary: #5c5c5c;
                    --input-bg: rgba(255, 255, 255, 0.7);
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

                /* 深色模式覆盖 */
                @media (prefers-color-scheme: dark) {
                    :root {
                        --panel-bg: rgba(32, 32, 32, 0.7);
                        --panel-border: rgba(255, 255, 255, 0.1);
                        --text-primary: #ffffff;
                        --text-secondary: #c7c7c7;
                        --input-bg: rgba(50, 50, 50, 0.7);
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
                    .winui-button.accent {
                        color: #1a1a1a !important;
                    }
                }

                #winui-draggable-panel {
                    position: fixed;
                    width: 340px;
                    z-index: 10000;
                    font-family: 'Segoe UI Variable', 'Segoe UI', system-ui, sans-serif;
                    pointer-events: none;
                }
                .winui-panel {
                    background: var(--panel-bg);
                    backdrop-filter: blur(25px) saturate(180%);
                    -webkit-backdrop-filter: blur(25px) saturate(180%);
                    border-radius: 12px;
                    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2), 0 0 0 1px var(--panel-border) inset;
                    border: 1px solid rgba(255, 255, 255, 0.05);
                    overflow: hidden;
                    pointer-events: auto;
                    transition: background 0.2s, box-shadow 0.3s;
                    color: var(--text-primary);
                }
                .winui-title {
                    font-size: 16px;
                    font-weight: 600;
                    padding: 16px 20px 8px;
                    color: var(--text-primary);
                    letter-spacing: -0.01em;
                    cursor: move;
                    user-select: none;
                }
                .winui-title:active {
                    cursor: grabbing;
                }
                .winui-content {
                    padding: 8px 20px 20px;
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
                    resize: vertical;
                }
                .winui-input-field input:focus,
                .winui-input-field textarea:focus {
                    border-color: var(--input-focus-border);
                    box-shadow: 0 0 0 3px var(--input-focus-shadow);
                    background: var(--input-bg);
                }
                .winui-button {
                    width: 100%;
                    padding: 10px 16px;
                    border-radius: 8px;
                    border: none;
                    font-weight: 500;
                    font-size: 14px;
                    cursor: pointer;
                    transition: all 0.25s cubic-bezier(0.2, 0.9, 0.4, 1);
                    background: var(--button-bg);
                    color: var(--text-primary);
                    border: 1px solid var(--button-border);
                    position: relative;
                    overflow: hidden;
                }
                .winui-button.accent {
                    background: var(--button-accent-bg);
                    border: none;
                    transform: scale(1);
                }
                .winui-button.accent:hover {
                    background: var(--button-accent-hover);
                }
                .winui-button.accent:active {
                    background: var(--button-accent-active);
                    transform: scale(0.97);
                }
                /* 成功状态动画 */
                .winui-button.accent.success {
                    background: #0e7a4d !important;
                    color: white !important;
                    transition: background 0.15s;
                }
                /* 响应式：小屏幕底部固定，禁用拖拽偏移 */
                @media (max-width: 700px) {
                    #winui-draggable-panel {
                        width: calc(100% - 20px) !important;
                        left: 10px !important;
                        right: 10px !important;
                        top: auto !important;
                        bottom: 10px !important;
                    }
                    .winui-title {
                        cursor: default;
                    }
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
        }

        _setupDragging() {
            const header = document.getElementById('drag-handle');
            if (!header) return;

            header.addEventListener('mousedown', (e) => {
                if (e.button !== 0) return;
                // 移动端不启用拖拽
                if (window.innerWidth <= 700) return;

                this.isDragging = true;
                this.startX = e.clientX;
                this.startY = e.clientY;
                const rect = this.panel.getBoundingClientRect();
                this.initialLeft = rect.left;
                this.initialTop = rect.top;
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
                    const left = parseInt(this.panel.style.left, 10);
                    const top = parseInt(this.panel.style.top, 10);
                    if (!isNaN(left) && !isNaN(top)) {
                        localStorage.setItem('winuiPanelLeft', left);
                        localStorage.setItem('winuiPanelTop', top);
                    }
                }
            };

            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        }

        _bindEvents() {
            this.btn.addEventListener('click', async () => {
                const descInput = document.getElementById('courseDescInput');
                const fbInput = document.getElementById('feedbackInput');
                const desc = descInput.value.trim();
                if (!desc) {
                    this.showErrorToast("请输入课程描述");
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
                this.btn.innerHTML = this.originalBtnText;
                this.timeoutId = null;
            }, 1800);
        }

        showErrorToast(message) {
            const toast = document.createElement('div');
            toast.textContent = message;
            toast.style.cssText = `
                position: fixed;
                bottom: 30px;
                right: 30px;
                background: var(--toast-bg-error);
                color: white;
                padding: 12px 20px;
                border-radius: 8px;
                font-size: 14px;
                z-index: 100000;
                box-shadow: 0 8px 20px rgba(0,0,0,0.3);
                font-family: 'Segoe UI Variable', system-ui, sans-serif;
                backdrop-filter: blur(10px);
                -webkit-backdrop-filter: blur(10px);
                border: 1px solid rgba(255,255,255,0.2);
                transition: opacity 0.3s;
                pointer-events: none;
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
                ["教师所在学院", data.teacherCollege],
                ["教师", data.teacher],
                ["课程", data.courseName],
                ["校区", data.campus],
                ["上课时间", data.time],
                ["课表教室", data.classroom],
                ["应到人数", data.studentCount],
                ["实到人数", data.studentCount],
                ["反馈内容", data.feedback],
                ["周次", data.week.toString()],
                ["信息员", data.officer],
                ["信息员所在学院", data.college]
            ];

            this.filler.fillAll(fields);

            const fullText = `学院：${data.college}\n信息员：${data.officer}\n校区：${data.campus}\n课表教室 ：${data.classroom}\n时间：${data.time}\n教师：${data.teacher}\n教师所在学院：${data.teacherCollege}\n课程：${data.courseName}\n人数：(${data.studentCount}/${data.studentCount})\n反馈内容：${data.feedback}`;

            const copied = await Utils.copyToClipboard(fullText);
            if (copied) {
                this.panel.showSuccessOnButton();
            } else {
                this.panel.showErrorToast("复制失败，请手动复制");
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