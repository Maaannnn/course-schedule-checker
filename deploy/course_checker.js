// 排课结果检查器
class CourseScheduleChecker {
    constructor() {
        this.scheduleData = null;
        this.results = [];
        
        // 艺术和体育课程的标识（可以根据实际情况调整）
        this.artSportsSubjects = ['美', '音', '体', '艺', '形'];
        
        this.setupEventListeners();
    }
    
    setupEventListeners() {
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');
        
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        
        // 拖拽功能
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFile(files[0]);
            }
        });
    }
    
    handleFileSelect(event) {
        const file = event.target.files[0];
        if (file) {
            this.handleFile(file);
        }
    }
    
    handleFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // 获取第一个工作表
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // 将工作表转换为JSON数组
                this.scheduleData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                this.checkSchedule();
            } catch (error) {
                this.displayError('文件读取失败：' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }
    
    // 解析课程信息
    parseCourseInfo(cellValue) {
        if (!cellValue || typeof cellValue !== 'string') {
            return null;
        }
        
        const trimmedValue = cellValue.trim();
        
        // 检查是否是体育课程
        if (trimmedValue.includes('体育')) {
            return {
                subject: '体',
                teacher: '体育老师',
                classroom: '',
                fullName: trimmedValue,
                isArtSports: true
            };
        }
        
        // 检查是否是艺术课程（包括复杂的艺术课程描述）
        if (trimmedValue.includes('艺术')) {
            // 解析艺术课程的班级信息
            const groupMatch = trimmedValue.match(/艺术(\d+)组/);
            const groupNumber = groupMatch ? groupMatch[1] : '';
            
            // 检查是否包含班级数量信息
            let classCount = 1;
            const classMatch = trimmedValue.match(/(两|三|四|五|六|七|八|九|十|[0-9]+)个?班级/);
            if (classMatch) {
                const countText = classMatch[1];
                if (countText === '两' || countText === '2') classCount = 2;
                else if (countText === '三' || countText === '3') classCount = 3;
                else if (countText === '四' || countText === '4') classCount = 4;
                else if (countText === '五' || countText === '5') classCount = 5;
                else if (!isNaN(parseInt(countText))) classCount = parseInt(countText);
            }
            
            return {
                subject: '艺',
                teacher: '艺术老师',
                classroom: '',
                fullName: trimmedValue,
                isArtSports: true,
                artGroupNumber: groupNumber,
                artClassCount: classCount
            };
        }
        
        const parts = trimmedValue.split(' ');
        if (parts.length < 2) {
            return null;
        }
        
        const subject = parts[0].charAt(0); // 课程科目（第一个字）
        const teacher = parts[1]; // 老师名字
        const classroom = parts.length > 2 ? parts[2] : ''; // 教室地点
        
        return {
            subject,
            teacher,
            classroom,
            fullName: trimmedValue,
            isArtSports: false
        };
    }
    
    // 解析Excel中的星期分组
    parseWeeklySchedule() {
        const weeklyData = {};
        let currentDay = null;
        let currentDayStartRow = -1;
        
        for (let rowIndex = 0; rowIndex < this.scheduleData.length; rowIndex++) {
            const row = this.scheduleData[rowIndex];
            if (row && row[1] && typeof row[1] === 'string' && row[1].includes('星期')) {
                // 找到星期标识行
                if (row[1].includes('星期一')) currentDay = '星期一';
                else if (row[1].includes('星期二')) currentDay = '星期二';
                else if (row[1].includes('星期三')) currentDay = '星期三';
                else if (row[1].includes('星期四')) currentDay = '星期四';
                else if (row[1].includes('星期五')) currentDay = '星期五';
                
                if (currentDay) {
                    currentDayStartRow = rowIndex;
                    weeklyData[currentDay] = {
                        startRow: rowIndex,
                        endRow: -1,
                        data: []
                    };
                }
            } else if (currentDay && currentDayStartRow !== -1) {
                // 如果遇到下一个星期或结束，设置上一天的结束行
                if (row && row[1] && typeof row[1] === 'string' && row[1].includes('星期') && !row[1].includes(currentDay)) {
                    weeklyData[currentDay].endRow = rowIndex - 1;
                } else if (rowIndex === this.scheduleData.length - 1) {
                    weeklyData[currentDay].endRow = rowIndex;
                }
            }
        }
        
        // 提取每天的课程数据
        for (const [day, info] of Object.entries(weeklyData)) {
            if (info.endRow === -1) {
                // 找到下一个星期开始的位置或文件结束
                const nextDayStart = this.findNextDayStart(info.startRow + 1);
                info.endRow = nextDayStart !== -1 ? nextDayStart - 1 : this.scheduleData.length - 1;
            }
            
            // 提取课程数据（跳过标题行）
            for (let i = info.startRow + 2; i <= info.endRow; i++) {
                if (this.scheduleData[i] && this.scheduleData[i][0] && 
                    !this.scheduleData[i][0].includes('年级') && 
                    !this.scheduleData[i][0].includes('班') &&
                    this.scheduleData[i][0] !== '班') {
                    info.data.push(this.scheduleData[i]);
                }
            }
        }
        
        return weeklyData;
    }
    
    findNextDayStart(fromRow) {
        for (let i = fromRow; i < this.scheduleData.length; i++) {
            const row = this.scheduleData[i];
            if (row && row[1] && typeof row[1] === 'string' && row[1].includes('星期')) {
                return i;
            }
        }
        return -1;
    }
    
    // 检查排课结果
    checkSchedule() {
        if (!this.scheduleData || this.scheduleData.length === 0) {
            this.displayError('没有找到有效的排课数据');
            return;
        }
        
        this.results = [];
        this.weeklyData = this.parseWeeklySchedule();
        
        // 检查1：每一行每个科目课是否只出现一次
        this.checkSubjectFrequencyPerRow();
        
        // 检查2：每一列（一天内）每个老师是否只有一节课
        this.checkTeacherFrequencyPerColumn();
        
        // 检查3：每一天的同一列不应该出现相同的教室
        this.checkClassroomConflictPerColumn();
        
        this.displayResults();
    }
    
    // 检查每一行每个科目课是否只出现一次
    checkSubjectFrequencyPerRow() {
        for (const [day, dayInfo] of Object.entries(this.weeklyData)) {
            for (let rowIndex = 0; rowIndex < dayInfo.data.length; rowIndex++) {
                const row = dayInfo.data[rowIndex];
                const subjectCount = {};
                const className = row[0]; // 班级名称
                
                for (let colIndex = 1; colIndex < row.length; colIndex++) {
                    const courseInfo = this.parseCourseInfo(row[colIndex]);
                    
                    if (courseInfo && !courseInfo.isArtSports) {
                        if (subjectCount[courseInfo.subject]) {
                            subjectCount[courseInfo.subject]++;
                        } else {
                            subjectCount[courseInfo.subject] = 1;
                        }
                    }
                }
                
                // 检查是否有重复的科目
                for (const [subject, count] of Object.entries(subjectCount)) {
                    if (count > 1) {
                        this.results.push({
                            type: 'error',
                            message: `${day} - 班级${className}：科目"${subject}"出现了${count}次，应该只出现一次`
                        });
                    }
                }
            }
        }
    }
    
    // 检查每一列每个老师是否只有一节课
    checkTeacherFrequencyPerColumn() {
        for (const [day, dayInfo] of Object.entries(this.weeklyData)) {
            if (dayInfo.data.length === 0) continue;
            
            const maxCols = Math.max(...dayInfo.data.map(row => row.length));
            
            for (let colIndex = 1; colIndex < maxCols; colIndex++) { // 从第2列开始（跳过班级名称列）
                const teacherCount = {};
                
                for (let rowIndex = 0; rowIndex < dayInfo.data.length; rowIndex++) {
                    if (dayInfo.data[rowIndex][colIndex]) {
                        const courseInfo = this.parseCourseInfo(dayInfo.data[rowIndex][colIndex]);
                        
                        if (courseInfo && !courseInfo.isArtSports) {
                            if (teacherCount[courseInfo.teacher]) {
                                teacherCount[courseInfo.teacher]++;
                            } else {
                                teacherCount[courseInfo.teacher] = 1;
                            }
                        }
                    }
                }
                
                // 检查是否有老师在同一时间段有多节课
                for (const [teacher, count] of Object.entries(teacherCount)) {
                    if (count > 1) {
                        this.results.push({
                            type: 'error',
                            message: `${day} - 第${colIndex}节课：老师"${teacher}"在同一时间段有${count}节课，应该只有一节课`
                        });
                    }
                }
            }
        }
    }
    
    // 检查每一天的同一列不应该出现相同的教室
    checkClassroomConflictPerColumn() {
        for (const [day, dayInfo] of Object.entries(this.weeklyData)) {
            if (dayInfo.data.length === 0) continue;
            
            const maxCols = Math.max(...dayInfo.data.map(row => row.length));
            
            for (let colIndex = 1; colIndex < maxCols; colIndex++) { // 从第2列开始（跳过班级名称列）
                const classroomCount = {};
                
                for (let rowIndex = 0; rowIndex < dayInfo.data.length; rowIndex++) {
                    if (dayInfo.data[rowIndex][colIndex]) {
                        const courseInfo = this.parseCourseInfo(dayInfo.data[rowIndex][colIndex]);
                        
                        if (courseInfo && courseInfo.classroom && !courseInfo.isArtSports) {
                            if (classroomCount[courseInfo.classroom]) {
                                classroomCount[courseInfo.classroom]++;
                            } else {
                                classroomCount[courseInfo.classroom] = 1;
                            }
                        }
                    }
                }
                
                // 检查是否有教室冲突
                for (const [classroom, count] of Object.entries(classroomCount)) {
                    if (count > 1) {
                        this.results.push({
                            type: 'error',
                            message: `${day} - 第${colIndex}节课：教室"${classroom}"在同一时间段被使用${count}次，存在冲突`
                        });
                    }
                }
            }
        }
    }
    
    // 显示检查结果
    displayResults() {
        const resultsDiv = document.getElementById('results');
        let html = '';
        
        if (this.results.length === 0) {
            html = `
                <div class="success">
                    <h3>✅ 排课检查完成</h3>
                    <p>恭喜！没有发现任何冲突或错误。排课结果符合要求。</p>
                </div>
            `;
        } else {
            html = `
                <div class="warning">
                    <h3>⚠️ 排课检查完成 - 发现 ${this.results.length} 个问题</h3>
                </div>
            `;
            
            // 按类别分组显示结果
            const groupedResults = this.groupResultsByCategory();
            
            for (const [category, results] of Object.entries(groupedResults)) {
                html += `
                    <div class="error">
                        <h4>【${category}】(${results.length}个问题):</h4>
                        <ul>
                `;
                
                results.forEach((result, index) => {
                    html += `<li><strong>问题 ${index + 1}:</strong> ${result.message}</li>`;
                });
                
                html += `
                        </ul>
                    </div>
                `;
            }
        }
        
        // 添加详细信息表格
        html += this.generateDetailTable();
        
        resultsDiv.innerHTML = html;
    }
    
    // 按类别分组结果
    groupResultsByCategory() {
        const grouped = {};
        
        this.results.forEach(result => {
            // 根据错误信息确定类别
            let category = '其他问题';
            if (result.message.includes('科目') && result.message.includes('出现了')) {
                category = '科目重复';
            } else if (result.message.includes('老师') && result.message.includes('在同一时间段有')) {
                category = '老师冲突';
            } else if (result.message.includes('教室') && result.message.includes('在同一时间段被使用')) {
                category = '教室冲突';
            }
            
            if (!grouped[category]) {
                grouped[category] = [];
            }
            grouped[category].push(result);
        });
        
        return grouped;
    }
    
    // 生成详细信息表格
    generateDetailTable() {
        if (!this.weeklyData || Object.keys(this.weeklyData).length === 0) {
            return '';
        }
        
        // 分析冲突信息，为冲突的课程分配颜色
        const conflictColors = this.analyzeConflicts();
        
        let tableHtml = `<h3>排课详情（按星期分组）</h3>`;
        
        const weekdays = ['星期一', '星期二', '星期三', '星期四', '星期五'];
        
        for (const day of weekdays) {
            if (!this.weeklyData[day] || this.weeklyData[day].data.length === 0) continue;
            
            tableHtml += `<h4>${day}</h4>`;
            tableHtml += `<table><thead><tr><th>班级</th>`;
            
            // 生成表头（节次）
            const maxCols = Math.max(...this.weeklyData[day].data.map(row => row.length));
            for (let i = 1; i < maxCols; i++) {
                tableHtml += `<th>第${i}节</th>`;
            }
            tableHtml += `</tr></thead><tbody>`;
            
            // 生成表格内容
            this.weeklyData[day].data.forEach((row, rowIndex) => {
                tableHtml += `<tr><td><strong>${row[0]}</strong></td>`;
                for (let colIndex = 1; colIndex < maxCols; colIndex++) {
                    const cellValue = row[colIndex] || '';
                    const courseInfo = this.parseCourseInfo(cellValue);
                    
                    let cellClass = '';
                    if (courseInfo && courseInfo.isArtSports) {
                        cellClass = 'style="background-color: #e3f2fd;"'; // 艺术体育课程高亮
                    } else if (courseInfo) {
                        // 检查是否有冲突
                        const conflictKey = `${day}-${colIndex}`;
                        if (conflictColors.teachers[conflictKey] && conflictColors.teachers[conflictKey][courseInfo.teacher]) {
                            cellClass = `style="background-color: ${conflictColors.teachers[conflictKey][courseInfo.teacher]}; border: 2px solid #d32f2f;"`;
                        } else if (conflictColors.classrooms[conflictKey] && conflictColors.classrooms[conflictKey][courseInfo.classroom]) {
                            cellClass = `style="background-color: ${conflictColors.classrooms[conflictKey][courseInfo.classroom]}; border: 2px solid #f57c00;"`;
                        }
                    }
                    
                    tableHtml += `<td ${cellClass}>${cellValue}</td>`;
                }
                tableHtml += `</tr>`;
            });
            
            tableHtml += `</tbody></table>`;
        }
        
        tableHtml += `
            <div style="margin-top: 20px;">
                <h4>图例说明：</h4>
                <div style="display: flex; flex-wrap: wrap; gap: 15px;">
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 20px; background-color: #e3f2fd; border: 1px solid #ccc; margin-right: 5px;"></div>
                        <span>艺术/体育课程</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 20px; background-color: #ffcdd2; border: 2px solid #d32f2f; margin-right: 5px;"></div>
                        <span>老师时间冲突</span>
                    </div>
                    <div style="display: flex; align-items: center;">
                        <div style="width: 20px; height: 20px; background-color: #ffe0b2; border: 2px solid #f57c00; margin-right: 5px;"></div>
                        <span>教室使用冲突</span>
                    </div>
                </div>
            </div>
        `;
        
        return tableHtml;
    }
    
    // 分析冲突并分配颜色
    analyzeConflicts() {
        const conflictColors = {
            teachers: {},
            classrooms: {}
        };
        
        // 定义冲突高亮颜色
        const teacherColors = ['#ffcdd2', '#f8bbd9', '#e1bee7', '#d1c4e9', '#c5cae9'];
        const classroomColors = ['#ffe0b2', '#ffccbc', '#d7ccc8', '#f0f4c3', '#dcedc8'];
        
        let teacherColorIndex = 0;
        let classroomColorIndex = 0;
        
        // 分析老师冲突
        for (const [day, dayInfo] of Object.entries(this.weeklyData)) {
            if (dayInfo.data.length === 0) continue;
            
            const maxCols = Math.max(...dayInfo.data.map(row => row.length));
            
            for (let colIndex = 1; colIndex < maxCols; colIndex++) {
                const teacherCount = {};
                const classroomCount = {};
                
                // 统计老师和教室使用情况
                for (let rowIndex = 0; rowIndex < dayInfo.data.length; rowIndex++) {
                    if (dayInfo.data[rowIndex][colIndex]) {
                        const courseInfo = this.parseCourseInfo(dayInfo.data[rowIndex][colIndex]);
                        
                        if (courseInfo && !courseInfo.isArtSports) {
                            // 统计老师
                            if (teacherCount[courseInfo.teacher]) {
                                teacherCount[courseInfo.teacher]++;
                            } else {
                                teacherCount[courseInfo.teacher] = 1;
                            }
                            
                            // 统计教室
                            if (courseInfo.classroom) {
                                if (classroomCount[courseInfo.classroom]) {
                                    classroomCount[courseInfo.classroom]++;
                                } else {
                                    classroomCount[courseInfo.classroom] = 1;
                                }
                            }
                        }
                    }
                }
                
                // 为有冲突的老师分配颜色
                const conflictKey = `${day}-${colIndex}`;
                for (const [teacher, count] of Object.entries(teacherCount)) {
                    if (count > 1) {
                        if (!conflictColors.teachers[conflictKey]) {
                            conflictColors.teachers[conflictKey] = {};
                        }
                        conflictColors.teachers[conflictKey][teacher] = teacherColors[teacherColorIndex % teacherColors.length];
                        teacherColorIndex++;
                    }
                }
                
                // 为有冲突的教室分配颜色
                for (const [classroom, count] of Object.entries(classroomCount)) {
                    if (count > 1) {
                        if (!conflictColors.classrooms[conflictKey]) {
                            conflictColors.classrooms[conflictKey] = {};
                        }
                        conflictColors.classrooms[conflictKey][classroom] = classroomColors[classroomColorIndex % classroomColors.length];
                        classroomColorIndex++;
                    }
                }
            }
        }
        
        return conflictColors;
    }
    
    displayError(message) {
        const resultsDiv = document.getElementById('results');
        resultsDiv.innerHTML = `
            <div class="error">
                <h3>❌ 错误</h3>
                <p>${message}</p>
            </div>
        `;
    }
}

// 初始化检查器
document.addEventListener('DOMContentLoaded', () => {
    new CourseScheduleChecker();
});

// 导出类以供其他脚本使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CourseScheduleChecker;
}
