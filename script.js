// 전역 변수
let subjects = {}; // 학기별 과목 정보
let students = []; // 학생 정보
let grades = {}; // 학생별 성적 정보
let zScoreSettings = {}; // 과목별 Z점수 설정

// 전역 변수로 이벤트 리스너 설정 여부 추적
let eventListenersSetup = false;

// DOM 로드 완료 후 실행
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM 로드 완료');
    initializeApp();
    loadData();
    setupEventListeners();
    console.log('초기화 완료');
});

// 페이지 로드 완료 후에도 한번 더 실행 (혹시 모를 경우)
window.addEventListener('load', function() {
    console.log('페이지 로드 완료');
    if (!eventListenersSetup) {
        console.log('이벤트 리스너 재설정');
        setupEventListeners();
    }
});

// 앱 초기화
function initializeApp() {
    // 학기별 과목 초기화
    const semesters = ['1-1', '1-2', '2-1', '2-2', '3-1'];
    semesters.forEach(semester => {
        subjects[semester] = [];
    });
    
    // 반 선택 옵션 업데이트
    updateClassSelect();
}

// 파일 업로드 생성 함수
function createFileUpload() {
    // 파일 입력을 완전히 새로 생성하여 이벤트 리스너 문제 해결
    const newFileInput = document.createElement('input');
    newFileInput.type = 'file';
    newFileInput.accept = '.xlsx,.xls';
    newFileInput.style.display = 'none';
    newFileInput.id = 'excel-file';
    newFileInput.addEventListener('change', handleExcelUpload);
    
    // body에 추가
    document.body.appendChild(newFileInput);
    
    // 파일 선택 대화상자 열기
    newFileInput.click();
    
    // 파일 선택 후 요소 제거 (약간의 지연 후)
    setTimeout(() => {
        if (newFileInput.parentNode) {
            newFileInput.parentNode.removeChild(newFileInput);
        }
    }, 1000);
}

// 이벤트 리스너 설정
function setupEventListeners() {
    console.log('이벤트 리스너 설정 시작');
    
    // 이미 설정되어 있으면 중복 설정 방지
    if (eventListenersSetup) {
        console.log('이벤트 리스너가 이미 설정되어 있습니다.');
        return;
    }
    
    // 탭 전환
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });

    // 과목 관리
    const addSubjectBtn = document.getElementById('add-subject');
    console.log('과목 추가 버튼:', addSubjectBtn);
    
    if (addSubjectBtn) {
        addSubjectBtn.addEventListener('click', addSubject);
        console.log('과목 추가 이벤트 리스너 설정 완료');
    } else {
        console.error('과목 추가 버튼을 찾을 수 없습니다!');
    }
    
    document.getElementById('subject-type').addEventListener('change', toggleZScoreSetup);
    
    // 학기 선택 변경 시 과목 목록 업데이트
    document.getElementById('semester-select').addEventListener('change', updateSubjectsDisplay);
    
    // 과목 조회 버튼
    document.getElementById('view-subjects').addEventListener('click', updateSubjectsDisplay);
    
    // 데이터 초기화
    document.getElementById('reset-data').addEventListener('click', resetAllData);

    // 통합 학생 및 성적 관리
    document.getElementById('download-template').addEventListener('click', downloadExcelTemplate);
    document.getElementById('upload-excel').addEventListener('click', createFileUpload);

    document.getElementById('save-all-data').addEventListener('click', saveAllData);
    document.getElementById('delete-selected').addEventListener('click', deleteSelectedStudents);
    
    // 학기 변경 시 테이블 자동 업데이트
    document.getElementById('bulk-semester').addEventListener('change', updateUnifiedTable);

    // 등급 확인
    document.getElementById('search-student').addEventListener('click', searchStudent);
    document.getElementById('export-student').addEventListener('click', exportStudentData);

    // 반별 보기
    document.getElementById('grade-select').addEventListener('change', updateClassSelect);
    document.getElementById('class-select').addEventListener('change', updateClassTable);
    document.getElementById('view-class').addEventListener('click', updateClassTable);
    document.getElementById('export-class').addEventListener('click', exportClassData);
    
    // 학기별 탭 전환
    document.querySelectorAll('.semester-tab').forEach(tab => {
        tab.addEventListener('click', () => switchSemesterTab(tab.dataset.semester));
    });

    // 모달 닫기
    document.querySelector('.close').addEventListener('click', closeModal);
    window.addEventListener('click', (e) => {
        if (e.target.classList.contains('modal')) {
            closeModal();
        }
    });
    
    // 이벤트 리스너 설정 완료 표시
    eventListenersSetup = true;
    console.log('이벤트 리스너 설정 완료');
}

// 탭 전환
function switchTab(tabName) {
    // 모든 탭 비활성화
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.classList.remove('active');
    });

    // 선택된 탭 활성화
    document.getElementById(tabName).classList.add('active');
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

    // 탭별 초기화
    if (tabName === 'setup') {
        updateSubjectsDisplay(); // 기본설정 탭으로 이동 시 과목 목록 업데이트
    } else if (tabName === 'class-view') {
        updateClassSelect();
    }
}

// 과목 추가
function addSubject() {
    console.log('과목 추가 함수 실행됨');
    
    const semester = document.getElementById('semester-select').value;
    const name = document.getElementById('subject-name').value.trim();
    const units = parseInt(document.getElementById('subject-units').value);
    const type = document.getElementById('subject-type').value;

    console.log('입력값:', { semester, name, units, type });

    if (!name || !units) {
        alert('과목명과 단위수를 입력해주세요.');
        return;
    }

    // Z점수 과목인 경우 평균, 표준편차 확인
    if (type === 'z-score') {
        const mean = parseFloat(document.getElementById('z-mean').value);
        const std = parseFloat(document.getElementById('z-std').value);
        
        if (!mean || !std) {
            alert('Z점수 과목의 경우 평균과 표준편차를 입력해주세요.');
            return;
        }
    }

    const subject = {
        id: Date.now(),
        name: name,
        units: units,
        type: type
    };

    console.log('추가할 과목:', subject);
    console.log('추가 전 과목 목록:', subjects[semester].map(s => s.name));

    subjects[semester].push(subject);
    
    console.log('추가 후 과목 목록:', subjects[semester].map(s => s.name));
    
    // Z점수 과목인 경우 설정도 함께 저장
    if (type === 'z-score') {
        const mean = parseFloat(document.getElementById('z-mean').value);
        const std = parseFloat(document.getElementById('z-std').value);
        
        if (!zScoreSettings[semester]) zScoreSettings[semester] = {};
        zScoreSettings[semester][subject.id] = {
            mean: mean,
            std: std
        };
    }
    
    console.log('추가 후 과목 목록:', subjects[semester]);
    
    updateSubjectsDisplay();
    updateUnifiedTable();
    saveData();

    // 폼 초기화
    document.getElementById('subject-name').value = '';
    document.getElementById('subject-units').value = '';
    document.getElementById('subject-type').value = 'direct';
    document.getElementById('z-mean').value = '';
    document.getElementById('z-std').value = '';
    document.getElementById('z-score-setup').style.display = 'none';
    
    console.log('과목 추가 완료');
}

// 과목 표시 업데이트
function updateSubjectsDisplay() {
    console.log('과목 표시 업데이트 함수 실행됨');
    
    // 기본설정 탭의 학기 선택에서 값을 가져오기
    const semester = document.getElementById('semester-select').value;
    const container = document.getElementById('subjects-container');
    
    console.log('현재 학기:', semester);
    console.log('과목 목록:', subjects[semester]);
    console.log('컨테이너:', container);
    
    if (!container) {
        console.log('과목 컨테이너를 찾을 수 없습니다.');
        return;
    }
    
    container.innerHTML = '';
    
    if (subjects[semester] && subjects[semester].length > 0) {
        subjects[semester].forEach((subject, index) => {
            console.log('과목 렌더링:', subject);
            
            const subjectDiv = document.createElement('div');
            subjectDiv.className = 'subject-item';
            subjectDiv.draggable = true;
            subjectDiv.dataset.subjectId = subject.id;
            subjectDiv.dataset.index = index;
            
            let zScoreInfo = '';
            if (subject.type === 'z-score') {
                const zScore = zScoreSettings[semester]?.[subject.id];
                if (zScore) {
                    zScoreInfo = `<span class="z-score-info">평균: ${zScore.mean}, 표준편차: ${zScore.std}</span>`;
                } else {
                    zScoreInfo = `<span class="z-score-info" style="color: #ff6b6b;">Z점수 설정 필요</span>`;
                }
            }
            
            subjectDiv.innerHTML = `
                <div class="subject-info">
                    <i class="fas fa-grip-vertical drag-handle" style="color: #999; margin-right: 10px; cursor: move;"></i>
                    <span class="subject-name">${subject.name}</span>
                    <span class="subject-units">${subject.units}단위</span>
                    <span class="subject-type">${subject.type === 'direct' ? '직접입력' : 'Z점수'}</span>
                    ${zScoreInfo}
                </div>
                <button class="btn btn-danger" onclick="removeSubject('${semester}', ${subject.id})">
                    <i class="fas fa-trash"></i> 삭제
                </button>
            `;
            container.appendChild(subjectDiv);
        });
        
        // 드래그 앤 드롭 이벤트 리스너 추가
        setupDragAndDrop(container, semester);
    } else {
        // 과목이 없을 때 메시지 표시
        container.innerHTML = '<div class="no-subjects">등록된 과목이 없습니다.</div>';
    }
    
    console.log('과목 표시 업데이트 완료');
}

// 드래그 앤 드롭 설정
function setupDragAndDrop(container, semester) {
    let draggedElement = null;
    let draggedIndex = null;

    // 드래그 시작
    container.addEventListener('dragstart', (e) => {
        if (e.target.classList.contains('subject-item')) {
            draggedElement = e.target;
            draggedIndex = parseInt(e.target.dataset.index);
            e.target.style.opacity = '0.5';
            e.dataTransfer.effectAllowed = 'move';
        }
    });

    // 드래그 종료
    container.addEventListener('dragend', (e) => {
        if (e.target.classList.contains('subject-item')) {
            e.target.style.opacity = '1';
        }
    });

    // 드래그 오버
    container.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        
        const target = e.target.closest('.subject-item');
        if (target && target !== draggedElement) {
            const targetIndex = parseInt(target.dataset.index);
            const rect = target.getBoundingClientRect();
            const midY = rect.top + rect.height / 2;
            
            if (e.clientY < midY) {
                target.style.borderTop = '2px solid #667eea';
                target.style.borderBottom = '';
            } else {
                target.style.borderTop = '';
                target.style.borderBottom = '2px solid #667eea';
            }
        }
    });

    // 드래그 리브
    container.addEventListener('dragleave', (e) => {
        const target = e.target.closest('.subject-item');
        if (target) {
            target.style.borderTop = '';
            target.style.borderBottom = '';
        }
    });

    // 드롭
    container.addEventListener('drop', (e) => {
        e.preventDefault();
        
        const target = e.target.closest('.subject-item');
        if (target && draggedElement && target !== draggedElement) {
            const targetIndex = parseInt(target.dataset.index);
            const rect = target.getBoundingClientRect();
            const midY = rect.top + rect.height / 2;
            
            // 드롭 위치에 따라 인덱스 조정
            let newIndex = targetIndex;
            if (e.clientY > midY) {
                newIndex = targetIndex + 1;
            }
            
            // 과목 순서 변경
            reorderSubjects(semester, draggedIndex, newIndex);
            
            // 시각적 효과 제거
            target.style.borderTop = '';
            target.style.borderBottom = '';
        }
    });
}

// 과목 순서 변경
function reorderSubjects(semester, fromIndex, toIndex) {
    const semesterSubjects = subjects[semester];
    if (!semesterSubjects || fromIndex === toIndex) return;
    
    console.log(`과목 순서 변경 시작: ${fromIndex} -> ${toIndex}`);
    console.log('변경 전 과목 순서:', semesterSubjects.map(s => s.name));
    
    // 배열에서 요소 제거
    const [movedSubject] = semesterSubjects.splice(fromIndex, 1);
    
    // 새로운 위치에 삽입
    semesterSubjects.splice(toIndex, 0, movedSubject);
    
    console.log('변경 후 과목 순서:', semesterSubjects.map(s => s.name));
    
    // 화면 업데이트
    updateSubjectsDisplay();
    updateUnifiedTable();
    saveData();
    
    console.log(`과목 순서 변경 완료: ${fromIndex} -> ${toIndex}`);
}

// 과목 삭제
function removeSubject(semester, subjectId) {
    if (confirm('이 과목을 삭제하시겠습니까?')) {
        subjects[semester] = subjects[semester].filter(s => s.id !== subjectId);
        
        // Z점수 설정도 함께 삭제
        if (zScoreSettings[semester] && zScoreSettings[semester][subjectId]) {
            delete zScoreSettings[semester][subjectId];
        }
        
        updateSubjectsDisplay();
        updateUnifiedTable();
        saveData();
    }
}

// Z점수 설정 토글
function toggleZScoreSetup() {
    const type = document.getElementById('subject-type').value;
    const zScoreSetup = document.getElementById('z-score-setup');
    
    if (type === 'z-score') {
        zScoreSetup.style.display = 'block';
    } else {
        zScoreSetup.style.display = 'none';
    }
}

// Z점수 설정 저장 (과목 추가 시 자동으로 저장됨)
function saveZScoreSettings() {
    // 이 함수는 더 이상 사용되지 않음
}

// 학생 추가
function addStudent() {
    const grade = document.getElementById('student-grade').value;
    const className = document.getElementById('student-class').value.trim();
    const name = document.getElementById('student-name').value.trim();
    const semester = document.getElementById('student-semester').value;
    
    // 학년과 학기를 조합하여 실제 학기 코드 생성
    const actualSemester = `${grade}-${semester.split('-')[1]}`;

    if (!className || !name) {
        alert('반과 학생명을 입력해주세요.');
        return;
    }

    const student = {
        id: Date.now(),
        grade: grade,
        class: className,
        name: name,
        semester: actualSemester
    };

    students.push(student);
    updateUnifiedTable();
    updateClassSelect();
    saveData();

    // 폼 초기화
    document.getElementById('student-class').value = '';
    document.getElementById('student-name').value = '';
    document.getElementById('student-semester').value = '1-1';
}



// 학생 삭제
function removeStudent(studentId) {
    if (confirm('이 학생을 삭제하시겠습니까?')) {
        students = students.filter(s => s.id !== studentId);
        delete grades[studentId];
        updateUnifiedTable();
        updateClassSelect();
        saveData();
    }
}

// 성적 입력 모달
function inputGrades(studentId) {
    const student = students.find(s => s.id === studentId);
    if (!student) return;

    const semester = student.semester;
    const semesterSubjects = subjects[semester] || [];
    
    if (semesterSubjects.length === 0) {
        alert('해당 학기에 등록된 과목이 없습니다. 먼저 과목을 등록해주세요.');
        return;
    }

    let modalContent = `
        <h3>${student.class}반 ${student.name} - ${getSemesterText(semester)} 성적 입력</h3>
        <form id="grade-input-form">
            <table class="grade-input-table">
                <thead>
                    <tr>
                        <th>과목명</th>
                        <th>단위수</th>
                        <th>입력값</th>
                        <th>유형</th>
                    </tr>
                </thead>
                <tbody>
    `;

    semesterSubjects.forEach(subject => {
        const currentGrade = grades[studentId]?.[semester]?.[subject.id] || '';
        modalContent += `
            <tr>
                <td>${subject.name}</td>
                <td>${subject.units}</td>
                <td>
                    <input type="number" 
                           id="grade-${subject.id}" 
                           placeholder="${subject.type === 'direct' ? '등급(1-9)' : '점수'}" 
                           value="${currentGrade}"
                           min="1" 
                           max="${subject.type === 'direct' ? '9' : '100'}"
                           step="${subject.type === 'direct' ? '1' : '0.1'}"
                           class="input-field">
                </td>
                <td>${subject.type === 'direct' ? '등급' : '점수'}</td>
            </tr>
        `;
    });

    modalContent += `
                </tbody>
            </table>
            <div style="display: flex; gap: 10px; justify-content: center; margin-top: 20px;">
                <button type="submit" class="btn btn-primary">저장</button>
                <button type="button" class="btn btn-secondary" onclick="closeModal()">취소</button>
            </div>
        </form>
    `;

    showModal(modalContent);

    // 폼 제출 이벤트
    document.getElementById('grade-input-form').addEventListener('submit', (e) => {
        e.preventDefault();
        saveGrades(studentId, semester, semesterSubjects);
    });
}

// 성적 저장
function saveGrades(studentId, semester, semesterSubjects) {
    if (!grades[studentId]) grades[studentId] = {};
    if (!grades[studentId][semester]) grades[studentId][semester] = {};

    let hasGrades = false;
    
    semesterSubjects.forEach(subject => {
        const gradeInput = document.getElementById(`grade-${subject.id}`);
        const gradeValue = gradeInput.value.trim();
        
        if (gradeValue) {
            hasGrades = true;
            let finalGrade;
            
            if (subject.type === 'z-score') {
                // Z점수 계산
                const score = parseFloat(gradeValue);
                const subjectZScore = zScoreSettings[semester]?.[subject.id];
                if (subjectZScore) {
                    const zScore = (score - subjectZScore.mean) / subjectZScore.std;
                    finalGrade = calculateGradeFromZScore(zScore);
                } else {
                    alert(`${subject.name} 과목의 Z점수 설정이 필요합니다.`);
                    return;
                }
            } else {
                // 직접 등급 입력
                finalGrade = parseInt(gradeValue);
            }
            
            grades[studentId][semester][subject.id] = {
                raw: gradeValue,
                grade: finalGrade,
                units: subject.units
            };
        }
    });

    if (hasGrades) {
        saveData();
        closeModal();
        alert('성적이 저장되었습니다.');
    } else {
        alert('최소 하나의 성적을 입력해주세요.');
    }
}

// Z점수로 등급 계산
function calculateGradeFromZScore(zScore) {
    // Z점수를 9등급으로 변환
    if (zScore >= 1.5) return 1;
    if (zScore >= 1.0) return 2;
    if (zScore >= 0.5) return 3;
    if (zScore >= 0.0) return 4;
    if (zScore >= -0.5) return 5;
    if (zScore >= -1.0) return 6;
    if (zScore >= -1.5) return 7;
    if (zScore >= -2.0) return 8;
    return 9;
}

// 학기별 총 등급 계산
function calculateSemesterGrade(studentId, semester) {
    const semesterGrades = grades[studentId]?.[semester];
    if (!semesterGrades) return null;

    let totalGradePoints = 0;
    let totalUnits = 0;

    Object.values(semesterGrades).forEach(gradeData => {
        totalGradePoints += gradeData.grade * gradeData.units;
        totalUnits += gradeData.units;
    });

    return totalUnits > 0 ? Math.round((totalGradePoints / totalUnits) * 100) / 100 : null;
}

// 대학별 등급 계산 (현재 입력된 데이터만 사용)
function calculateUniversityGrades(studentId) {
    const semesters = ['1-1', '1-2', '2-1', '2-2', '3-1'];
    const semesterGrades = {};
    
    // 각 학기별 등급 계산 (데이터가 있는 경우만)
    semesters.forEach(semester => {
        const grade = calculateSemesterGrade(studentId, semester);
        if (grade !== null) {
            semesterGrades[semester] = grade;
        }
    });

    // A대학: 1학년 20%, 2학년 30%, 3학년 1학기 50%
    let universityA = null;
    const availableGrades = Object.values(semesterGrades).filter(g => g !== null);
    
    if (availableGrades.length > 0) {
        if (semesterGrades['1-1'] && semesterGrades['1-2'] && semesterGrades['2-1'] && semesterGrades['2-2'] && semesterGrades['3-1']) {
            // 모든 학기 데이터가 있는 경우
            const grade1 = (semesterGrades['1-1'] + semesterGrades['1-2']) / 2;
            const grade2 = (semesterGrades['2-1'] + semesterGrades['2-2']) / 2;
            const grade3 = semesterGrades['3-1'];
            universityA = (grade1 * 0.2 + grade2 * 0.3 + grade3 * 0.5);
        } else if (semesterGrades['3-1']) {
            // 3학년 1학기만 있는 경우
            universityA = semesterGrades['3-1'];
        } else {
            // 가중 평균 계산 (가용한 데이터만 사용)
            const totalWeight = availableGrades.length;
            universityA = availableGrades.reduce((sum, grade) => sum + grade, 0) / totalWeight;
        }
    }

    // B대학: 최고학기 50%, 3학년 1학기 50%
    let universityB = null;
    if (semesterGrades['3-1']) {
        const grades1to2 = [semesterGrades['1-1'], semesterGrades['1-2'], semesterGrades['2-1'], semesterGrades['2-2']].filter(g => g !== null);
        if (grades1to2.length > 0) {
            const bestGrade = Math.min(...grades1to2);
            universityB = (bestGrade * 0.5 + semesterGrades['3-1'] * 0.5);
        } else {
            universityB = semesterGrades['3-1'];
        }
    } else if (availableGrades.length > 0) {
        // 3학년 1학기가 없는 경우, 가용한 데이터로 계산
        universityB = Math.min(...availableGrades);
    }

    return {
        universityA: universityA ? Math.round(universityA * 100) / 100 : null,
        universityB: universityB ? Math.round(universityB * 100) / 100 : null
    };
}

// 학생 검색
function searchStudent() {
    const searchGrade = document.getElementById('search-grade').value;
    const searchClass = document.getElementById('search-class').value.trim();
    const searchName = document.getElementById('search-name').value.trim();

    if (!searchClass && !searchName) {
        alert('반 또는 학생명을 입력해주세요.');
        return;
    }

    let filteredStudents = students;
    
    // 학년으로 필터링
    if (searchGrade) {
        filteredStudents = filteredStudents.filter(s => s.grade == searchGrade || getGradeFromSemester(s.semester) == searchGrade);
    }

    const student = filteredStudents.find(s => 
        s.class === searchClass && s.name === searchName
    );

    if (!student) {
        alert('해당 학생을 찾을 수 없습니다.');
        return;
    }

    displayStudentGrades(student);
}



// 데이터 초기화
function resetAllData() {
    if (confirm('모든 데이터를 초기화하시겠습니까? 이 작업은 되돌릴 수 없습니다.')) {
        subjects = {};
        students = [];
        grades = {};
        zScoreSettings = {};
        
        // 학기별 과목 초기화
        const semesters = ['1-1', '1-2', '2-1', '2-2', '3-1'];
        semesters.forEach(semester => {
            subjects[semester] = [];
        });
        
        // 화면 업데이트
        updateUnifiedTable();
        updateClassSelect();
        
        // 로컬 스토리지 초기화
        localStorage.removeItem('gradeCalculatorData');
        
        alert('모든 데이터가 초기화되었습니다.');
    }
}

// 통합 테이블 초기화 및 업데이트
function updateUnifiedTable() {
    const semester = document.getElementById('bulk-semester').value;
    const semesterSubjects = subjects[semester] || [];
    
    if (semesterSubjects.length === 0) {
        // 과목이 없으면 빈 테이블 표시
        createEmptyUnifiedTable();
        return;
    }
    
    // 해당 학기의 학생들 가져오기 (기존 데이터에서)
    const semesterStudents = students.filter(s => s.semester === semester);
    createUnifiedTable(semester, semesterSubjects, semesterStudents);
}

// 빈 통합 테이블 생성
function createEmptyUnifiedTable() {
    const header = document.getElementById('unified-table-header');
    const body = document.getElementById('unified-table-body');
    
    header.innerHTML = '<tr><th><input type="checkbox" id="select-all" onchange="toggleSelectAll(this)"></th><th>반</th><th>학생명</th><th colspan="3">과목을 먼저 등록해주세요</th></tr>';
    body.innerHTML = '';
}

// 통합 테이블 생성
function createUnifiedTable(semester, subjects, existingStudents) {
    const header = document.getElementById('unified-table-header');
    const body = document.getElementById('unified-table-body');
    
    // 헤더 생성
    let headerHTML = '<tr><th><input type="checkbox" id="select-all" onchange="toggleSelectAll(this)"></th><th>반</th><th>학생명</th>';
    subjects.forEach(subject => {
        headerHTML += `<th>${subject.name}<br><small>(${subject.type === 'direct' ? '등급' : '점수'})</small></th>`;
    });
    headerHTML += '</tr>';
    header.innerHTML = headerHTML;
    
    // 바디 생성
    body.innerHTML = '';
    
    // 기존 학생들 표시
    existingStudents.forEach(student => {
        addStudentRowToTable(student, subjects, semester, true);
    });
    
    // 새 학생 추가 행
    addNewStudentRow(subjects, semester);
    
    // 칸 추가 버튼 추가
    addAddRowButton(body, subjects, semester);
    
    // 자동 붙여넣기 기능 활성화
    setupAutoPaste(body);
    
    // 전역 붙여넣기 이벤트 리스너 추가 (테이블 외부에서도 작동)
    document.removeEventListener('paste', handleGlobalPaste);
    document.addEventListener('paste', handleGlobalPaste);
}

// 학생 행을 테이블에 추가
function addStudentRowToTable(student, subjects, semester, isExisting = false) {
    const body = document.getElementById('unified-table-body');
    const row = document.createElement('tr');
    
    if (isExisting) {
        row.className = 'existing-student';
    }
    
    // 체크박스 추가
    row.innerHTML = `<td><input type="checkbox" class="student-checkbox" data-student-id="${isExisting ? student.id : ''}"></td>`;
    
    // 반과 학생명 (기존 학생도 수정 가능하도록)
    if (isExisting) {
        row.innerHTML += `
            <td><input type="text" class="class-input existing-class" value="${student.class}"></td>
            <td><input type="text" class="name-input existing-name" value="${student.name}"></td>
        `;
    } else {
        row.innerHTML += `
            <td><input type="text" class="class-input" placeholder="반"></td>
            <td><input type="text" class="name-input" placeholder="학생명"></td>
        `;
    }
    
    // 과목별 성적 입력
    subjects.forEach(subject => {
        if (isExisting) {
            let currentGrade = '';
            const gradeData = grades[student.id]?.[semester]?.[subject.id]?.raw;
            if (gradeData && !isNaN(parseFloat(gradeData))) {
                currentGrade = gradeData.toString();
            }
            
            row.innerHTML += `
                <td>
                    <input type="number" 
                           class="grade-input existing-grade"
                           data-student-id="${student.id}" 
                           data-subject-id="${subject.id}"
                           value="${currentGrade}"
                           min="${subject.type === 'direct' ? '1' : '0'}" 
                           max="${subject.type === 'direct' ? '9' : '100'}"
                           step="${subject.type === 'direct' ? '1' : '0.1'}">
                </td>
            `;
        } else {
            row.innerHTML += `
                <td>
                    <input type="number" 
                           class="grade-input new-grade"
                           data-subject-id="${subject.id}"
                           min="${subject.type === 'direct' ? '1' : '0'}" 
                           max="${subject.type === 'direct' ? '9' : '100'}"
                           step="${subject.type === 'direct' ? '1' : '0.1'}">
                </td>
            `;
        }
    });
    
    // 새 학생 행에는 삭제 버튼 추가
    if (!isExisting) {
        row.innerHTML += '<td><button class="btn btn-danger btn-sm" onclick="removeNewStudentRow(this)"><i class="fas fa-trash"></i></button></td>';
    }
    
    body.appendChild(row);
}

// 새 학생 행 추가
function addNewStudentRow(subjects, semester) {
    const body = document.getElementById('unified-table-body');
    const row = document.createElement('tr');
    row.className = 'new-student';
    
    // 체크박스 추가
    row.innerHTML = `<td><input type="checkbox" class="student-checkbox"></td>`;
    
    // 반과 학생명 입력
    row.innerHTML += `
        <td><input type="text" class="class-input" placeholder="반"></td>
        <td><input type="text" class="name-input" placeholder="학생명"></td>
    `;
    
    // 과목별 성적 입력
    subjects.forEach(subject => {
        row.innerHTML += `
            <td>
                <input type="number" 
                       class="grade-input new-grade"
                       data-subject-id="${subject.id}"
                       min="${subject.type === 'direct' ? '1' : '0'}" 
                       max="${subject.type === 'direct' ? '9' : '100'}"
                       step="${subject.type === 'direct' ? '1' : '0.1'}">
            </td>
        `;
    });
    
    // 삭제 버튼
    row.innerHTML += '<td><button class="btn btn-danger btn-sm" onclick="removeNewStudentRow(this)"><i class="fas fa-trash"></i></button></td>';
    
    body.appendChild(row);
}

// 칸 추가 버튼 생성
function addAddRowButton(body, subjects, semester) {
    // 기존 칸 추가 버튼이 있으면 제거
    const existingAddRowDiv = body.querySelector('.add-row-container');
    if (existingAddRowDiv) {
        existingAddRowDiv.remove();
    }
    
    const addRowDiv = document.createElement('div');
    addRowDiv.className = 'add-row-container';
    addRowDiv.innerHTML = `
        <button class="btn btn-primary add-row-btn" onclick="addNewStudentRowFromButton('${semester}')">
            <i class="fas fa-plus"></i> 칸 추가
        </button>
        <small style="color: #666; margin-left: 10px;">Ctrl+V로 엑셀 데이터 붙여넣기 가능 (콘솔 로그 제외)</small>
    `;
    body.appendChild(addRowDiv);
}

// 칸 추가 버튼으로 새 행 추가
function addNewStudentRowFromButton(semester) {
    const semesterSubjects = subjects[semester] || [];
    const body = document.getElementById('unified-table-body');
    
    // 새 학생 행 추가
    addNewStudentRow(semesterSubjects, semester);
    
    // 새로운 칸 추가 버튼 추가 (중복 방지 로직 포함)
    addAddRowButton(body, semesterSubjects, semester);
}

// 자동 붙여넣기 설정
function setupAutoPaste(body) {
    // 기존 이벤트 리스너 제거
    body.removeEventListener('paste', handleAutoPaste);
    // 새 이벤트 리스너 추가
    body.addEventListener('paste', handleAutoPaste);
    console.log('붙여넣기 이벤트 리스너 설정됨');
}

// 자동 붙여넣기 처리
function handleAutoPaste(e) {
    console.log('붙여넣기 이벤트 발생');
    e.preventDefault();
    
    const clipboardData = e.clipboardData || window.clipboardData;
    const pastedData = clipboardData.getData('text');
    
    console.log('붙여넣기 데이터:', JSON.stringify(pastedData));
    console.log('붙여넣기 데이터 길이:', pastedData.length);
    console.log('붙여넣기 데이터의 각 문자 코드:', Array.from(pastedData).map(c => c.charCodeAt(0)));
    
    if (!pastedData) {
        console.log('붙여넣기 데이터가 없음');
        return;
    }
    
    // 줄바꿈 문자 정규화 (Windows, Mac, Linux 호환)
    const normalizedData = pastedData.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const rows = normalizedData.split('\n').filter(row => row.trim());
    
    console.log('정규화된 데이터:', JSON.stringify(normalizedData));
    console.log('정규화된 데이터의 각 문자 코드:', Array.from(normalizedData).map(c => c.charCodeAt(0)));
    console.log('분할된 행들:', rows.map((row, i) => `행${i}: ${JSON.stringify(row)}`));
    console.log('행 개수:', rows.length);
    
    const table = document.getElementById('unified-data-table');
    const tableBody = table.querySelector('tbody');
    const semester = document.getElementById('bulk-semester').value;
    const semesterSubjects = subjects[semester] || [];
    
    console.log('붙여넣기 정보:', {
        rowsLength: rows.length,
        semester,
        subjectsLength: semesterSubjects.length,
        originalDataLength: pastedData.length,
        normalizedDataLength: normalizedData.length
    });
    
    // 콘솔 로그가 아닌 실제 엑셀 데이터인지 확인
    const isConsoleLog = pastedData.includes('script.js:') || 
                        pastedData.includes('행 생성 완료:') || 
                        pastedData.includes('성공적으로 추가됨') ||
                        pastedData.includes('<tr class=');
    
    if (isConsoleLog) {
        alert('콘솔 로그가 아닌 실제 엑셀 데이터를 복사해서 붙여넣어주세요.\n\n올바른 방법:\n1. 엑셀에서 데이터를 선택\n2. Ctrl+C로 복사\n3. 테이블에서 Ctrl+V로 붙여넣기');
        return;
    }
    
    if (rows.length < 2 || semesterSubjects.length === 0) {
        alert('올바른 데이터 형식이 아닙니다. 반, 학생명, 과목점수 형식으로 입력해주세요.');
        return;
    }
    
    // 기존 새 학생 행들과 칸 추가 버튼 제거
    const existingNewRows = tableBody.querySelectorAll('.new-student');
    existingNewRows.forEach(row => row.remove());
    
    const existingAddRowDiv = tableBody.querySelector('.add-row-container');
    if (existingAddRowDiv) {
        existingAddRowDiv.remove();
    }
    
    let processedRows = 0;
    
    // 모든 행을 데이터로 처리 (헤더가 없는 경우)
    console.log('=== 행 처리 시작 ===');
    console.log('전체 행 개수:', rows.length);
    console.log('처리할 행 개수:', rows.length);
    
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        console.log(`\n=== 행 ${i} 처리 시작 ===`);
        console.log(`처리 중인 행 ${i}:`, JSON.stringify(row));
        console.log(`행 ${i} 길이:`, row.length);
        
        const cells = row.split('\t');
        console.log(`행 ${i}의 셀들:`, cells);
        console.log(`행 ${i}의 셀 개수:`, cells.length);
        
        if (cells.length < 3) {
            console.log(`행 ${i} 건너뛰기: 셀 개수 부족 (${cells.length})`);
            continue;
        }
        
        const className = cells[0].trim();
        const studentName = cells[1].trim();
        
        console.log(`행 ${i} 파싱 결과:`, { className, studentName });
        console.log(`반 길이: ${className.length}, 학생명 길이: ${studentName.length}`);
        
        if (!className || !studentName) {
            console.log(`행 ${i} 건너뛰기: 반 또는 학생명이 비어있음`);
            continue;
        }
        
        console.log('행 데이터 처리:', { className, studentName, cells });
        
        // 새 행 생성
        const newRow = createNewStudentRowFromPaste(className, studentName, semesterSubjects, cells);
        if (newRow) {
            tableBody.appendChild(newRow);
            processedRows++;
            console.log(`행 ${i} 성공적으로 추가됨 (총 ${processedRows}개 처리됨)`);
        } else {
            console.log(`행 ${i} 추가 실패`);
        }
    }
    
    console.log('=== 행 처리 완료 ===');
    console.log('최종 처리된 행 수:', processedRows);
    
    // 칸 추가 버튼 다시 추가 (한 번만)
    addAddRowButton(tableBody, semesterSubjects, semester);
    
    // 붙여넣기된 데이터를 자동으로 저장하고 기존 학생 행으로 변환
    if (processedRows > 0) {
        savePastedData();
        // 테이블을 새로고침하여 기존 학생 행으로 표시
        setTimeout(() => {
            updateUnifiedTable();
        }, 100);
    }
    
    // 성공 메시지
    showQuickMessage(`엑셀 데이터가 성공적으로 붙여넣어졌습니다. (처리된 행: ${processedRows}개, 전체 행: ${rows.length}개)`);
}

// 전역 붙여넣기 핸들러
function handleGlobalPaste(e) {
    // 학생관리 탭이 활성화되어 있고, 과목이 등록되어 있는 경우에만 처리
    const activeTab = document.querySelector('.tab-content.active');
    if (activeTab && activeTab.id === 'students') {
        const semester = document.getElementById('bulk-semester').value;
        const semesterSubjects = subjects[semester] || [];
        
        if (semesterSubjects.length > 0) {
            console.log('전역 붙여넣기 이벤트 처리');
            handleAutoPaste(e);
        }
    }
}

// 붙여넣기된 데이터 자동 저장
function savePastedData() {
    const semester = document.getElementById('bulk-semester').value;
    const semesterSubjects = subjects[semester] || [];
    
    if (semesterSubjects.length === 0) {
        console.log('저장할 과목이 없습니다.');
        return;
    }
    
    let newStudentsCount = 0;
    let updatedGradesCount = 0;
    
    // 새로 붙여넣기된 학생들 처리
    const newStudentRows = document.querySelectorAll('.new-student');
    newStudentRows.forEach(row => {
        const classInput = row.querySelector('.class-input');
        const nameInput = row.querySelector('.name-input');
        const gradeInputs = row.querySelectorAll('.grade-input');
        
        if (!classInput || !nameInput) return;
        
        const className = classInput.value.trim();
        const studentName = nameInput.value.trim();
        
        if (!className || !studentName) return;
        
        // 새 학생 생성
        const student = {
            id: Date.now() + Math.random(),
            grade: semester.split('-')[0],
            class: className,
            name: studentName,
            semester: semester
        };
        
        students.push(student);
        newStudentsCount++;
        
        // 성적 저장
        if (!grades[student.id]) grades[student.id] = {};
        if (!grades[student.id][semester]) grades[student.id][semester] = {};
        
        gradeInputs.forEach((input, index) => {
            const value = input.value.trim();
            // 숫자 값만 허용하고 유효성 검사
            if (value && !isNaN(parseFloat(value)) && index < semesterSubjects.length) {
                const subject = semesterSubjects[index];
                let finalGrade = null;
                
                if (subject.type === 'z-score') {
                    const score = parseFloat(value);
                    if (!isNaN(score)) {
                        const subjectZScore = zScoreSettings[semester]?.[subject.id];
                        if (subjectZScore) {
                            const zScore = (score - subjectZScore.mean) / subjectZScore.std;
                            finalGrade = calculateGradeFromZScore(zScore);
                        }
                    }
                } else {
                    const grade = parseInt(value);
                    if (!isNaN(grade)) {
                        finalGrade = grade;
                    }
                }
                
                grades[student.id][semester][subject.id] = {
                    raw: value,
                    grade: finalGrade,
                    units: subject.units
                };
                updatedGradesCount++;
            }
        });
    });
    
    // 데이터 저장
    saveData();
    
    // 화면 업데이트
    updateClassSelect();
    
    console.log(`붙여넣기 데이터 자동 저장 완료: 새 학생 ${newStudentsCount}명, 성적 ${updatedGradesCount}개`);
    
    // 성공 메시지 업데이트
    showQuickMessage(`데이터가 자동으로 저장되었습니다. (새 학생: ${newStudentsCount}명, 성적: ${updatedGradesCount}개)`);
}

// 빠른 메시지 표시 함수
function showQuickMessage(message) {
    // 기존 메시지 제거
    const existingMessage = document.querySelector('.quick-message');
    if (existingMessage) {
        existingMessage.remove();
    }
    
    // 새 메시지 생성
    const messageDiv = document.createElement('div');
    messageDiv.className = 'quick-message';
    messageDiv.textContent = message;
    messageDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #28a745;
        color: white;
        padding: 10px 15px;
        border-radius: 5px;
        z-index: 10000;
        font-size: 14px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.2);
    `;
    
    document.body.appendChild(messageDiv);
    
    // 3초 후 자동 제거
    setTimeout(() => {
        if (messageDiv.parentNode) {
            messageDiv.remove();
        }
    }, 3000);
}

// 빈 학생 행 생성
function createEmptyStudentRow(subjects, semester) {
    const row = document.createElement('tr');
    row.className = 'new-student';
    
    // 체크박스 추가
    row.innerHTML = `<td><input type="checkbox" class="student-checkbox"></td>`;
    
    // 반과 학생명 입력 (빈 값)
    row.innerHTML += `
        <td><input type="text" class="class-input" placeholder="반"></td>
        <td><input type="text" class="name-input" placeholder="학생명"></td>
    `;
    
    // 과목별 성적 입력 (빈 값)
    subjects.forEach(subject => {
        row.innerHTML += `
            <td>
                <input type="number" 
                       class="grade-input new-grade"
                       data-subject-id="${subject.id}"
                       placeholder=""
                       min="${subject.type === 'direct' ? '1' : '0'}" 
                       max="${subject.type === 'direct' ? '9' : '100'}"
                       step="${subject.type === 'direct' ? '1' : '0.1'}">
            </td>
        `;
    });
    
    // 삭제 버튼
    row.innerHTML += '<td><button class="btn btn-danger btn-sm" onclick="removeNewStudentRow(this)"><i class="fas fa-trash"></i></button></td>';
    
    return row;
}

// 행에 데이터 채우기
function fillRowWithData(row, className, studentName, subjects, cells) {
    // 반과 학생명 채우기
    const classInput = row.querySelector('.class-input');
    const nameInput = row.querySelector('.name-input');
    
    if (classInput) classInput.value = className;
    if (nameInput) nameInput.value = studentName;
    
    // 과목별 성적 채우기
    const gradeInputs = row.querySelectorAll('.grade-input');
    subjects.forEach((subject, index) => {
        if (gradeInputs[index]) {
            let value = '';
            if (cells[index + 2]) {
                const cellValue = cells[index + 2].toString().trim();
                // 숫자만 허용하고 안전한 값으로 처리
                if (cellValue && !isNaN(parseFloat(cellValue))) {
                    value = cellValue;
                }
            }
            gradeInputs[index].value = value;
        }
    });
}

// 새 학생 행 삭제
function removeNewStudentRow(button) {
    button.closest('tr').remove();
}

// 전체 선택/해제 토글
function toggleSelectAll(checkbox) {
    const studentCheckboxes = document.querySelectorAll('.student-checkbox');
    studentCheckboxes.forEach(cb => {
        cb.checked = checkbox.checked;
    });
}

// 선택된 학생들 삭제
function deleteSelectedStudents() {
    const selectedCheckboxes = document.querySelectorAll('.student-checkbox:checked');
    
    if (selectedCheckboxes.length === 0) {
        alert('삭제할 학생을 선택해주세요.');
        return;
    }
    
    if (!confirm(`선택된 ${selectedCheckboxes.length}명의 학생을 삭제하시겠습니까?`)) {
        return;
    }
    
    let deletedCount = 0;
    
    selectedCheckboxes.forEach(checkbox => {
        const row = checkbox.closest('tr');
        const studentId = checkbox.dataset.studentId;
        
        if (studentId) {
            // 기존 학생 삭제
            const studentIndex = students.findIndex(s => s.id == studentId);
            if (studentIndex !== -1) {
                students.splice(studentIndex, 1);
                delete grades[studentId];
                deletedCount++;
            }
        }
        
        // 행 삭제
        row.remove();
    });
    
    // 전체 선택 체크박스 해제
    const selectAllCheckbox = document.getElementById('select-all');
    if (selectAllCheckbox) {
        selectAllCheckbox.checked = false;
    }
    
    // 데이터 저장 및 화면 업데이트
    saveData();
    updateClassSelect();
    
    alert(`${deletedCount}명의 학생이 삭제되었습니다.`);
}

// 페이지 로드 시 테이블 초기화
document.addEventListener('DOMContentLoaded', function() {
    // 기존 초기화 코드...
    setTimeout(() => {
        updateUnifiedTable();
    }, 100);
});



// 붙여넣기용 새 학생 행 생성
function createNewStudentRowFromPaste(className, studentName, subjects, cells) {
    console.log('createNewStudentRowFromPaste 호출:', { className, studentName, subjectsLength: subjects.length, cells });
    
    const row = document.createElement('tr');
    row.className = 'new-student';
    
    // 체크박스 추가
    row.innerHTML = `<td><input type="checkbox" class="student-checkbox"></td>`;
    
    // 반과 학생명 입력
    row.innerHTML += `
        <td><input type="text" class="class-input" value="${className}"></td>
        <td><input type="text" class="name-input" value="${studentName}"></td>
    `;
    
    // 과목별 성적 입력
    subjects.forEach((subject, index) => {
        let value = '';
        if (cells[index + 2]) {
            const cellValue = cells[index + 2].toString().trim();
            // 숫자만 허용하고 안전한 값으로 처리
            if (cellValue && !isNaN(parseFloat(cellValue))) {
                value = cellValue;
            }
        }
        
        console.log(`과목 ${index} (${subject.name}): 셀값=${cells[index + 2]}, 최종값=${value}`);
        
        row.innerHTML += `
            <td>
                <input type="number" 
                       class="grade-input new-grade"
                       data-subject-id="${subject.id}"
                       value="${value}"
                       min="${subject.type === 'direct' ? '1' : '0'}" 
                       max="${subject.type === 'direct' ? '9' : '100'}"
                       step="${subject.type === 'direct' ? '1' : '0.1'}">
            </td>
        `;
    });
    
    // 삭제 버튼
    row.innerHTML += '<td><button class="btn btn-danger btn-sm" onclick="removeNewStudentRow(this)"><i class="fas fa-trash"></i></button></td>';
    
    console.log('행 생성 완료:', row);
    return row;
}

// 모든 데이터 저장 (학생 정보 + 성적)
function saveAllData() {
    const semester = document.getElementById('bulk-semester').value;
    const semesterSubjects = subjects[semester] || [];
    
    if (semesterSubjects.length === 0) {
        alert('저장할 과목이 없습니다. 먼저 과목을 등록해주세요.');
        return;
    }
    
    let newStudentsCount = 0;
    let updatedGradesCount = 0;
    
    // 새 학생들 처리
    const newStudentRows = document.querySelectorAll('.new-student');
    newStudentRows.forEach(row => {
        const classInput = row.querySelector('.class-input');
        const nameInput = row.querySelector('.name-input');
        const gradeInputs = row.querySelectorAll('.grade-input');
        
        if (!classInput || !nameInput) return;
        
        const className = classInput.value.trim();
        const studentName = nameInput.value.trim();
        
        if (!className || !studentName) return;
        
        // 새 학생 생성
        const student = {
            id: Date.now() + Math.random(),
            grade: semester.split('-')[0],
            class: className,
            name: studentName,
            semester: semester
        };
        
        students.push(student);
        newStudentsCount++;
        
        // 성적 저장
        if (!grades[student.id]) grades[student.id] = {};
        if (!grades[student.id][semester]) grades[student.id][semester] = {};
        
        gradeInputs.forEach((input, index) => {
            const value = input.value.trim();
            // 숫자 값만 허용하고 유효성 검사
            if (value && !isNaN(parseFloat(value)) && index < semesterSubjects.length) {
                const subject = semesterSubjects[index];
                let finalGrade = null;
                
                if (subject.type === 'z-score') {
                    const score = parseFloat(value);
                    if (!isNaN(score)) {
                        const subjectZScore = zScoreSettings[semester]?.[subject.id];
                        if (subjectZScore) {
                            const zScore = (score - subjectZScore.mean) / subjectZScore.std;
                            finalGrade = calculateGradeFromZScore(zScore);
                        }
                    }
                } else {
                    const grade = parseInt(value);
                    if (!isNaN(grade)) {
                        finalGrade = grade;
                    }
                }
                
                grades[student.id][semester][subject.id] = {
                    raw: value,
                    grade: finalGrade,
                    units: subject.units
                };
                updatedGradesCount++;
            }
        });
    });
    
    // 기존 학생들의 정보 및 성적 업데이트
    const existingRows = document.querySelectorAll('.existing-student');
    existingRows.forEach(row => {
        const classInput = row.querySelector('.existing-class');
        const nameInput = row.querySelector('.existing-name');
        const gradeInputs = row.querySelectorAll('.existing-grade');
        
        if (!classInput || !nameInput) return;
        
        const className = classInput.value.trim();
        const studentName = nameInput.value.trim();
        
        if (!className || !studentName) return;
        
        // 기존 학생 찾기
        const studentId = gradeInputs[0]?.dataset.studentId;
        if (!studentId) return;
        
        const student = students.find(s => s.id == studentId);
        if (student) {
            // 학생 정보 업데이트
            student.class = className;
            student.name = studentName;
            
            // 성적 업데이트
            gradeInputs.forEach(input => {
                const subjectId = input.dataset.subjectId;
                const value = input.value.trim();
                
                // 숫자 값만 허용하고 유효성 검사
                if (!value || isNaN(parseFloat(value))) return;
                
                if (!grades[studentId]) grades[studentId] = {};
                if (!grades[studentId][semester]) grades[studentId][semester] = {};
                
                const subject = subjects[semester]?.find(s => s.id == subjectId);
                if (subject) {
                    let finalGrade = null;
                    
                    if (subject.type === 'z-score') {
                        const score = parseFloat(value);
                        if (!isNaN(score)) {
                            const subjectZScore = zScoreSettings[semester]?.[subjectId];
                            if (subjectZScore) {
                                const zScore = (score - subjectZScore.mean) / subjectZScore.std;
                                finalGrade = calculateGradeFromZScore(zScore);
                            }
                        }
                    } else {
                        const grade = parseInt(value);
                        if (!isNaN(grade)) {
                            finalGrade = grade;
                        }
                    }
                    
                    grades[studentId][semester][subjectId] = {
                        raw: value,
                        grade: finalGrade,
                        units: subject.units
                    };
                    updatedGradesCount++;
                }
            });
        }
    });
    
    // 데이터 저장 및 화면 업데이트
    saveData();
    updateClassSelect();
    
    // 새 학생 행들 제거하고 테이블 새로고침
    newStudentRows.forEach(row => row.remove());
    updateUnifiedTable();
    
    alert(`저장 완료!\n새로 추가된 학생: ${newStudentsCount}명\n업데이트된 성적: ${updatedGradesCount}개`);
}

// 학생 등급 표시
function displayStudentGrades(student) {
    const displayDiv = document.getElementById('student-grade-display');
    const studentNameDiv = document.getElementById('display-student-name');
    const semesterGradesDiv = document.getElementById('semester-grades');
    
    studentNameDiv.textContent = `${student.class}반 ${student.name}`;
    
    // 학기별 등급 표시 (모든 과목 상세 포함)
    const semesters = ['1-1', '1-2', '2-1', '2-2', '3-1'];
    let semesterHTML = '<h4>학기별 등급</h4>';
    
    semesters.forEach(semester => {
        const semesterSubjects = subjects[semester] || [];
        const semesterGrades = grades[student.id]?.[semester] || {};
        const totalGrade = calculateSemesterGrade(student.id, semester);
        
        semesterHTML += `
            <div class="semester-grade-item">
                <div class="semester-header">
                    <strong>${getSemesterText(semester)}</strong>
                    <span class="total-grade">총등급: ${totalGrade ? totalGrade.toFixed(2) : '-'}</span>
                </div>
                <div class="subject-grades">
        `;
        
        if (semesterSubjects.length > 0) {
            semesterSubjects.forEach(subject => {
                const gradeData = semesterGrades[subject.id];
                const gradeDisplay = gradeData ? gradeData.grade : '-';
                const rawDisplay = gradeData ? gradeData.raw : '-';
                
                semesterHTML += `
                    <div class="subject-grade-detail">
                        <span class="subject-name">${subject.name} (${subject.units}단위)</span>
                        <span class="grade-value">${gradeDisplay}</span>
                        ${subject.type === 'z-score' ? `<span class="raw-score">원점수: ${rawDisplay}</span>` : ''}
                    </div>
                `;
            });
        } else {
            semesterHTML += '<div class="no-subjects">등록된 과목이 없습니다.</div>';
        }
        
        semesterHTML += `
                </div>
            </div>
        `;
    });
    
    semesterGradesDiv.innerHTML = semesterHTML;
    
    // 대학별 등급 계산 및 표시
    const universityGrades = calculateUniversityGrades(student.id);
    document.getElementById('university-a-grade').textContent = 
        universityGrades.universityA ? universityGrades.universityA.toFixed(2) : '-';
    document.getElementById('university-b-grade').textContent = 
        universityGrades.universityB ? universityGrades.universityB.toFixed(2) : '-';
    
    displayDiv.style.display = 'block';
}

// 반 선택 업데이트
function updateClassSelect() {
    const gradeSelect = document.getElementById('grade-select');
    const classSelect = document.getElementById('class-select');
    const selectedGrade = gradeSelect ? gradeSelect.value : '';
    
    // 선택된 학년의 학생들만 필터링
    let filteredStudents = students;
    if (selectedGrade) {
        filteredStudents = students.filter(s => s.grade == selectedGrade || getGradeFromSemester(s.semester) == selectedGrade);
    }
    
    const classes = [...new Set(filteredStudents.map(s => s.class))];
    
    classSelect.innerHTML = '<option value="">반 선택</option>';
    classes.forEach(className => {
        const option = document.createElement('option');
        option.value = className;
        option.textContent = `${className}반`;
        classSelect.appendChild(option);
    });
}

// 반별 테이블 업데이트
function updateClassTable() {
    const selectedClass = document.getElementById('class-select').value;
    if (!selectedClass) return;

    // 현재 선택된 학기 탭 가져오기
    const activeTab = document.querySelector('.semester-tab.active');
    const currentSemester = activeTab ? activeTab.dataset.semester : '1-1';
    
    updateClassTableForSemester(selectedClass, currentSemester);
}

// 학기별 탭 전환
function switchSemesterTab(semester) {
    // 모든 탭 비활성화
    document.querySelectorAll('.semester-tab').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // 선택된 탭 활성화
    document.querySelector(`[data-semester="${semester}"]`).classList.add('active');
    
    // 선택된 반이 있으면 테이블 업데이트
    const selectedClass = document.getElementById('class-select').value;
    if (selectedClass) {
        updateClassTableForSemester(selectedClass, semester);
    }
}

// 특정 학기의 반별 테이블 업데이트
function updateClassTableForSemester(selectedClass, semester) {
    const classStudents = students.filter(s => s.class === selectedClass);
    const semesterSubjects = subjects[semester] || [];
    
    // 헤더 업데이트
    const header = document.getElementById('class-table-header');
    let headerHTML = '<tr><th>학생명</th>';
    
    semesterSubjects.forEach(subject => {
        headerHTML += `<th>${subject.name} (${subject.units}단위)</th>`;
    });
    
    headerHTML += '<th>총등급</th></tr>';
    header.innerHTML = headerHTML;
    
    // 바디 업데이트
    const tbody = document.getElementById('class-table-body');
    tbody.innerHTML = '';
    
    classStudents.forEach(student => {
        const row = document.createElement('tr');
        let rowHTML = `<td>${student.name}</td>`;
        
        let totalGradePoints = 0;
        let totalUnits = 0;
        
        semesterSubjects.forEach(subject => {
            const gradeData = grades[student.id]?.[semester]?.[subject.id];
            if (gradeData && gradeData.grade !== null) {
                rowHTML += `<td>${gradeData.grade}</td>`;
                totalGradePoints += gradeData.grade * gradeData.units;
                totalUnits += gradeData.units;
            } else {
                rowHTML += '<td class="empty-grade">-</td>';
            }
        });
        
        const totalGrade = totalUnits > 0 ? (totalGradePoints / totalUnits).toFixed(2) : '-';
        rowHTML += `<td>${totalGrade}</td>`;
        
        row.innerHTML = rowHTML;
        tbody.appendChild(row);
    });
}

// 엑셀 템플릿 다운로드
function downloadExcelTemplate() {
    // 현재 활성화된 탭에 따라 적절한 학기 선택기에서 값을 가져오기
    let semester;
    const activeTab = document.querySelector('.tab-content.active');
    
    if (activeTab && activeTab.id === 'setup') {
        // 기본설정 탭인 경우
        semester = document.getElementById('semester-select').value;
    } else {
        // 학생관리 탭인 경우
        semester = document.getElementById('bulk-semester').value;
    }
    
    const semesterSubjects = subjects[semester] || [];
    
    if (semesterSubjects.length === 0) {
        alert('먼저 과목을 등록해주세요.');
        return;
    }
    
    console.log('엑셀 다운로드 - 학기:', semester);
    console.log('엑셀 다운로드 - 과목 순서:', semesterSubjects.map(s => s.name));
    
    // 헤더 생성 - 과목 순서대로 (입력 순서 유지)
    const header = ['반', '학생명'];
    semesterSubjects.forEach(subject => {
        header.push(`${subject.name} (${subject.type === 'direct' ? '등급' : '점수'})`);
    });
    
    console.log('엑셀 헤더:', header);
    
    // 예시 데이터
    const template = [header];
    
    // 샘플 데이터 추가 (하나의 예시만)
    if (semesterSubjects.length > 0) {
        const sampleData = ['1', '홍길동'];
        semesterSubjects.forEach(s => {
            if (s.type === 'direct') {
                sampleData.push('3'); // 등급 예시
            } else {
                sampleData.push('85.5'); // 점수 예시
            }
        });
        template.push(sampleData);
    }

    const ws = XLSX.utils.aoa_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `${getSemesterText(semester)}_성적입력양식`);
    
    XLSX.writeFile(wb, `${getSemesterText(semester)}_성적입력양식.xlsx`);
}

// 엑셀 파일 업로드 처리
function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // 헤더 가져오기
            const header = jsonData[0];
            if (!header || header.length < 3) {
                alert('올바른 엑셀 양식이 아닙니다.');
                return;
            }

            // 현재 선택된 학기
            const semester = document.getElementById('bulk-semester').value;
            const semesterSubjects = subjects[semester] || [];
            
            if (semesterSubjects.length === 0) {
                alert('먼저 과목을 등록해주세요.');
                return;
            }

            // 헤더에서 과목명 매핑 - 과목 순서대로 매핑
            const subjectMapping = {};
            
            // 헤더에서 과목명을 순서대로 찾아서 매핑 (엑셀의 열 순서와 과목 순서 매칭)
            semesterSubjects.forEach((subject, subjectIndex) => {
                const expectedHeader = `${subject.name} (${subject.type === 'direct' ? '등급' : '점수'})`;
                const headerIndex = header.indexOf(expectedHeader);
                if (headerIndex !== -1) {
                    subjectMapping[subjectIndex] = headerIndex;
                }
            });

            // 기존 테이블 내용 초기화
            const tableBody = document.getElementById('unified-table-body');
            tableBody.innerHTML = '';

            let addedCount = 0;

            // 데이터 처리 (헤더 제외)
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row.length < 2) continue;

                const className = row[0];
                const studentName = row[1];
                
                // 반과 학생명이 모두 있는 경우만 처리
                if (className && studentName) {
                    // 새 학생 행 추가
                    const newRow = createNewStudentRowFromExcel(className, studentName, semesterSubjects, row, subjectMapping);
                    tableBody.appendChild(newRow);
                    addedCount++;
                }
            }

            // 화면 업데이트
            updateClassSelect();
            
            alert(`엑셀 업로드 완료!\n추가된 학생: ${addedCount}명\n\n이제 "모든 데이터 저장" 버튼을 클릭하여 데이터를 저장하세요.`);

        } catch (error) {
            console.error('엑셀 파일 처리 오류:', error);
            alert('엑셀 파일 처리 중 오류가 발생했습니다: ' + error.message);
        }
    };
    
    reader.onerror = function() {
        alert('파일 읽기 중 오류가 발생했습니다.');
    };
    
    reader.readAsArrayBuffer(file);
}

// 엑셀 업로드용 새 학생 행 생성
function createNewStudentRowFromExcel(className, studentName, subjects, row, subjectMapping) {
    const tableRow = document.createElement('tr');
    tableRow.className = 'new-student';
    
    // 체크박스 추가
    tableRow.innerHTML = `<td><input type="checkbox" class="student-checkbox"></td>`;
    
    // 반과 학생명 입력
    tableRow.innerHTML += `
        <td><input type="text" class="class-input" value="${className}"></td>
        <td><input type="text" class="name-input" value="${studentName}"></td>
    `;
    
    // 과목별 성적 입력
    subjects.forEach((subject, index) => {
        const headerIndex = subjectMapping[index];
        let value = '';
        
        if (headerIndex !== undefined && row[headerIndex] !== undefined) {
            const cellValue = row[headerIndex];
            // 숫자만 허용하고 안전한 값으로 처리
            if (cellValue !== null && cellValue !== undefined) {
                const strValue = cellValue.toString().trim();
                if (strValue && !isNaN(parseFloat(strValue))) {
                    value = strValue;
                }
            }
        }
        
        tableRow.innerHTML += `
            <td>
                <input type="number" 
                       class="grade-input new-grade"
                       data-subject-id="${subject.id}"
                       value="${value}"
                       min="${subject.type === 'direct' ? '1' : '0'}" 
                       max="${subject.type === 'direct' ? '9' : '100'}"
                       step="${subject.type === 'direct' ? '1' : '0.1'}">
            </td>
        `;
    });
    
    // 삭제 버튼
    tableRow.innerHTML += '<td><button class="btn btn-danger btn-sm" onclick="removeNewStudentRow(this)"><i class="fas fa-trash"></i></button></td>';
    
    return tableRow;
}

// 학생별 데이터 내보내기
function exportStudentData() {
    const searchClass = document.getElementById('search-class').value.trim();
    const searchName = document.getElementById('search-name').value.trim();
    
    if (!searchClass || !searchName) {
        alert('먼저 학생을 검색해주세요.');
        return;
    }

    const student = students.find(s => s.class === searchClass && s.name === searchName);
    if (!student) return;

    const data = [];
    const semesters = ['1-1', '1-2', '2-1', '2-2', '3-1'];
    
    // 학생 기본 정보를 첫 번째 행에 한 번만 표시
    data.push([
        student.class,
        student.name,
        '전체 학기',
        '학생 정보',
        '-',
        '-',
        '-'
    ]);
    
    // 구분선 추가
    data.push(['', '', '', '', '', '', '']);
    
    semesters.forEach(semester => {
        const semesterGrades = grades[student.id]?.[semester];
        if (semesterGrades && Object.keys(semesterGrades).length > 0) {
            // 학기별 구분선 추가
            data.push(['', '', getSemesterText(semester), '', '', '', '']);
            
            Object.entries(semesterGrades).forEach(([subjectId, gradeData]) => {
                const subject = subjects[semester]?.find(s => s.id == subjectId);
                if (subject) {
                    data.push([
                        '',  // 반 (비워둠)
                        '',  // 학생명 (비워둠)
                        '',  // 학기 (비워둠)
                        subject.name,
                        subject.units,
                        gradeData.raw,
                        gradeData.grade
                    ]);
                }
            });
            
            // 학기별 총등급 추가
            const totalGrade = calculateSemesterGrade(student.id, semester);
            if (totalGrade !== null) {
                data.push([
                    '',  // 반 (비워둠)
                    '',  // 학생명 (비워둠)
                    '',  // 학기 (비워둠)
                    '총등급',
                    '-',
                    '-',
                    totalGrade.toFixed(2)
                ]);
            }
            
            // 학기 간 구분선 추가
            data.push(['', '', '', '', '', '', '']);
        }
    });

    const ws = XLSX.utils.aoa_to_sheet([
        ['반', '학생명', '학기', '과목명', '단위수', '원점수', '등급'],
        ...data
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '학생성적');
    
    XLSX.writeFile(wb, `${student.class}반_${student.name}_성적.xlsx`);
}

// 반별 데이터 내보내기
function exportClassData() {
    const selectedClass = document.getElementById('class-select').value;
    if (!selectedClass) {
        alert('반을 선택해주세요.');
        return;
    }

    const classStudents = students.filter(s => s.class === selectedClass);
    const semesters = ['1-1', '1-2', '2-1', '2-2', '3-1'];
    
    // 모든 학기를 한 시트에 표시
    const wb = XLSX.utils.book_new();
    
    // 전체 데이터 배열
    const allData = [];
    
    semesters.forEach(semester => {
        const semesterSubjects = subjects[semester] || [];
        if (semesterSubjects.length === 0) return;
        
        // 학기 구분선 추가
        allData.push([`=== ${getSemesterText(semester)} ===`]);
        
        // 헤더 생성
        const header = ['학생명'];
        semesterSubjects.forEach(subject => {
            header.push(`${subject.name} (${subject.units}단위)`);
        });
        header.push('총등급');
        allData.push(header);
        
        // 학생별 데이터
        classStudents.forEach(student => {
            const row = [student.name];
            let totalGradePoints = 0;
            let totalUnits = 0;
            
            semesterSubjects.forEach(subject => {
                const gradeData = grades[student.id]?.[semester]?.[subject.id];
                if (gradeData && gradeData.grade !== null) {
                    row.push(gradeData.grade);
                    totalGradePoints += gradeData.grade * gradeData.units;
                    totalUnits += gradeData.units;
                } else {
                    row.push('-');
                }
            });
            
            const totalGrade = totalUnits > 0 ? (totalGradePoints / totalUnits).toFixed(2) : '-';
            row.push(totalGrade);
            allData.push(row);
        });
        
        // 학기 간 빈 줄 추가
        allData.push([]);
    });
    
    // 대학별 등급 추가
    allData.push(['=== 대학별 등급 ===']);
    allData.push(['학생명', 'A대학 등급', 'B대학 등급']);
    classStudents.forEach(student => {
        const universityGrades = calculateUniversityGrades(student.id);
        allData.push([
            student.name,
            universityGrades.universityA ? universityGrades.universityA.toFixed(2) : '-',
            universityGrades.universityB ? universityGrades.universityB.toFixed(2) : '-'
        ]);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(allData);
    XLSX.utils.book_append_sheet(wb, ws, '전체성적');
    
    XLSX.writeFile(wb, `${selectedClass}반_성적.xlsx`);
}

// 모달 표시
function showModal(content) {
    document.getElementById('modal-body').innerHTML = content;
    document.getElementById('modal').style.display = 'block';
}

// 모달 닫기
function closeModal() {
    document.getElementById('modal').style.display = 'none';
}

// 학기 텍스트 변환
function getSemesterText(semester) {
    const semesterMap = {
        '1-1': '1학년 1학기',
        '1-2': '1학년 2학기',
        '2-1': '2학년 1학기',
        '2-2': '2학년 2학기',
        '3-1': '3학년 1학기'
    };
    return semesterMap[semester] || semester;
}

// 학기에서 학년 추출
function getGradeFromSemester(semester) {
    return semester.split('-')[0];
}

// 데이터 저장
function saveData() {
    const data = {
        subjects: subjects,
        students: students,
        grades: grades,
        zScoreSettings: zScoreSettings
    };
    localStorage.setItem('gradeCalculatorData', JSON.stringify(data));
    console.log('데이터 저장됨 - 과목 순서:', Object.keys(subjects).map(semester => ({
        semester,
        subjects: subjects[semester]?.map(s => s.name) || []
    })));
}

// 데이터 로드
function loadData() {
    const savedData = localStorage.getItem('gradeCalculatorData');
    if (savedData) {
        try {
            const data = JSON.parse(savedData);
            subjects = data.subjects || {};
            students = data.students || [];
            grades = data.grades || {};
            zScoreSettings = data.zScoreSettings || {};
            
            // 학기별 과목 초기화 (새로운 학기 추가 시)
            const semesters = ['1-1', '1-2', '2-1', '2-2', '3-1'];
            semesters.forEach(semester => {
                if (!subjects[semester]) subjects[semester] = [];
            });
            
            console.log('데이터 로드됨 - 과목 순서:', Object.keys(subjects).map(semester => ({
                semester,
                subjects: subjects[semester]?.map(s => s.name) || []
            })));
            
            updateUnifiedTable();
            updateClassSelect();
            updateSubjectsDisplay(); // 기본설정 탭의 과목 목록도 업데이트
        } catch (error) {
            console.error('데이터 로드 중 오류:', error);
        }
    }
}


