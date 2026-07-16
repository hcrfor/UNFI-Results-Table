/**
 * 도시산림자원조사 결과표 자동 생성기 - 핵심 로직 스크립트
 * 시니어 풀스택 개발자 코딩 멘토 버전
 */

document.addEventListener('DOMContentLoaded', () => {
    // UI 요소 캐싱
    const dropZone = document.getElementById('drop-zone');
    const folderInput = document.getElementById('folder-input');
    const selectFolderBtn = document.getElementById('select-folder-btn');
    
    const dashboardSection = document.getElementById('dashboard-section');
    const resultsSection = document.getElementById('results-section');
    
    const countExcelEl = document.getElementById('count-excel');
    const countPhotosEl = document.getElementById('count-photos');
    const countMatchedEl = document.getElementById('count-matched');
    const countUnmatchedEl = document.getElementById('count-unmatched');
    
    const processStatusEl = document.getElementById('process-status');
    const progressBar = document.getElementById('analysis-progress');
    const progressText = document.getElementById('progress-text');
    const processingFileEl = document.getElementById('processing-file');
    
    const logAccordion = document.getElementById('log-accordion');
    const logHeader = document.getElementById('log-header');
    const logBody = document.getElementById('log-body');
    const logList = document.getElementById('log-list');
    const logCountEl = document.getElementById('log-count');
    
    const tableBody = document.getElementById('table-body');
    const btnReset = document.getElementById('btn-reset');
    const btnDownload = document.getElementById('btn-download');
    const themeToggle = document.getElementById('theme-toggle');
    const themeIcon = document.getElementById('theme-icon');

    // 분석 데이터를 담을 메인 상태 객체
    // key: 표본점번호 (파일명/폴더명)
    let surveyState = {};
    let logLogs = [];
    let topFolderName = '도시산림자원조사';

    // 1. 테마(다크/라이트 모드) 변경 기능
    // 기본 테마 설정 체크
    const savedTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);

    themeToggle.addEventListener('click', () => {
        const currentTheme = document.documentElement.getAttribute('data-theme');
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('theme', newTheme);
        updateThemeIcon(newTheme);
        addLog('system', `화면 테마를 ${newTheme === 'dark' ? '다크 모드' : '라이트 모드'}로 변경했습니다.`);
    });

    function updateThemeIcon(theme) {
        if (theme === 'dark') {
            themeIcon.setAttribute('data-lucide', 'sun');
            themeIcon.style.color = '#fbbf24';
        } else {
            themeIcon.setAttribute('data-lucide', 'moon');
            themeIcon.style.color = '';
        }
        if (window.lucide) lucide.createIcons();
    }

    // 2. 로그 기록 헬퍼 함수
    function addLog(type, message) {
        const time = new Date().toLocaleTimeString('ko-KR', { hour12: false });
        let typeBadge = '';
        let itemClass = '';

        if (type === 'error') {
            typeBadge = '[오류]';
            itemClass = 'text-red';
        } else if (type === 'warning') {
            typeBadge = '[경고]';
            itemClass = 'text-amber';
        } else if (type === 'success') {
            typeBadge = '[성공]';
            itemClass = 'text-green';
        } else {
            typeBadge = '[정보]';
            itemClass = '';
        }

        const logText = `[${time}] ${typeBadge} ${message}`;
        logLogs.push({ time, type, message });

        const logItem = document.createElement('div');
        logItem.className = `log-item ${itemClass}`;
        logItem.innerHTML = `<span class="log-time">[${time}]</span> <span class="log-msg">${typeBadge} ${message}</span>`;
        logList.appendChild(logItem);
        
        // 로그 수 업데이트 및 아코디언이 켜져있다면 자동 스크롤
        logCountEl.textContent = logLogs.length;
        logBody.scrollTop = logBody.scrollHeight;
    }

    // 로그 아코디언 열고 닫기
    logHeader.addEventListener('click', () => {
        logAccordion.classList.toggle('active');
        logBody.classList.toggle('hidden');
    });

    // 3. 파일 선택 단추 클릭 연동
    selectFolderBtn.addEventListener('click', () => {
        folderInput.click();
    });

    folderInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            processUploadedFiles(Array.from(e.target.files));
        }
    });

    // 4. 드래그 앤 드롭 이벤트 핸들러
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    // 4.5. 드래그앤드롭 시 브라우저가 누락시키는 webkitRelativePath 상대경로를 재귀적으로 복원하는 비동기 헬퍼 함수
    async function scanFilesFromEntry(entry, path = '') {
        let files = [];
        if (entry.isFile) {
            const file = await new Promise((resolve, reject) => {
                entry.file(resolve, reject);
            });
            // 드래그앤드롭 시 유실되는 relativePath를 직접 복구해 주입합니다!
            // 예: "test_dataset/야장폴더/410977119997.xlsx"
            const customPath = path ? `${path}/${file.name}` : file.name;
            Object.defineProperty(file, 'webkitRelativePath', {
                value: customPath,
                writable: false,
                configurable: true
            });
            files.push(file);
        } else if (entry.isDirectory) {
            const dirReader = entry.createReader();
            
            // readEntries가 페이징(기본 100개 제한)되어 올 수 있으므로 완전 스캔 루프 구현
            const getEntries = () => new Promise((resolve) => {
                const allEntries = [];
                const read = () => {
                    dirReader.readEntries((results) => {
                        if (results.length === 0) {
                            resolve(allEntries);
                        } else {
                            allEntries.push(...results);
                            read();
                        }
                    }, () => resolve(allEntries));
                };
                read();
            });

            const entries = await getEntries();
            for (const childEntry of entries) {
                // entry.name을 경로에 포함시켜 재귀적으로 상대 경로 생성
                const childFiles = await scanFilesFromEntry(childEntry, path ? `${path}/${entry.name}` : entry.name);
                files.push(...childFiles);
            }
        }
        return files;
    }

    dropZone.addEventListener('drop', async (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        
        const items = e.dataTransfer.items;
        if (items && items.length > 0) {
            // 분석 상태 초기화 피드백
            progressBar.style.width = '5%';
            progressText.textContent = '드롭된 폴더 스캔 준비 중...';
            processingFileEl.textContent = '폴더 구조를 재구성하는 중입니다.';
            
            try {
                const entry = items[0].webkitGetAsEntry();
                if (entry) {
                    topFolderName = entry.name;
                    addLog('info', `드래그앤드롭 폴더명 감지: ${topFolderName}`);
                    
                    // 재귀 스캔을 통한 상대경로(webkitRelativePath) 복원 파일 목록 획득
                    const files = await scanFilesFromEntry(entry);
                    addLog('info', `하위 파일 스캔 완료 (총 ${files.length}개 파일). 분석을 시작합니다.`);
                    
                    if (files.length === 0) {
                        alert('스캔된 파일이 없습니다. 폴더 구조를 확인해주세요.');
                        return;
                    }
                    
                    processUploadedFiles(files);
                } else {
                    const files = Array.from(e.dataTransfer.files);
                    processUploadedFiles(files);
                }
            } catch (err) {
                console.error(err);
                addLog('error', `드래그 폴더 스캔 중 실패: ${err.message}`);
                alert(`폴더 구조를 읽는 도중 오류가 발생했습니다: ${err.message}`);
            }
        }
    });

    // 5. 폴더 데이터 정리 및 파일 매핑 시작
    async function processUploadedFiles(files) {
        // 상태 초기화
        surveyState = {};
        logLogs = [];
        logList.innerHTML = '';
        tableBody.innerHTML = '';
        
        // 최상위 폴더명 추출
        if (files.length > 0 && files[0].webkitRelativePath) {
            const pathParts = files[0].webkitRelativePath.split('/');
            if (pathParts.length > 0 && pathParts[0] !== '') {
                topFolderName = pathParts[0];
            }
        }
        
        addLog('info', `총 ${files.length}개의 파일을 감지하여 분석을 시작합니다. (최상위 폴더: ${topFolderName})`);

        // 야장폴더와 사진폴더를 분리하여 정리
        let excelCount = 0;
        let photoFileCount = 0;
        
        // 경로 세그먼트 분석을 통한 분류
        files.forEach(file => {
            const path = file.webkitRelativePath;
            const pathParts = path.split('/');
            
            // 야장 엑셀 파일 검출
            // 규칙: 확장자가 xlsx 이고 경로상 '야장' 혹은 'field'가 포함된 경우
            const isExcel = file.name.endsWith('.xlsx') && !file.name.startsWith('~$'); // 임시파일 제외
            const isInExcelFolder = pathParts.some(part => part.includes('야장') || part.toLowerCase().includes('field'));
            
            // 사진 파일 검출
            // 규칙: 이미지 확장자 이고 경로상 '사진' 혹은 'photo'가 포함된 경우
            const isPhoto = /\.(jpg|jpeg|png|gif|webp)$/i.test(file.name);
            const isInPhotoFolder = pathParts.some(part => part.includes('사진') || part.toLowerCase().includes('photo') || part.toLowerCase().includes('image'));

            if (isExcel && isInExcelFolder) {
                // 표본점번호는 파일명(확장자 제외)으로 추출
                const sampleNo = file.name.substring(0, file.name.lastIndexOf('.')).trim();
                if (!surveyState[sampleNo]) {
                    initSurveyItem(sampleNo);
                }
                surveyState[sampleNo].excelFile = file;
                surveyState[sampleNo].excelStatus = 'O';
                excelCount++;
            } else if (isPhoto && isInPhotoFolder) {
                // 사진의 경우, 파일이 담겨 있는 바로 상위 폴더명이 표본점번호가 됨
                // 예: Root/사진폴더/410977119997/P_2026.jpg -> parts[parts.length - 2] = '410977119997'
                if (pathParts.length >= 3) {
                    const sampleNo = pathParts[pathParts.length - 2].trim();
                    if (!surveyState[sampleNo]) {
                        initSurveyItem(sampleNo);
                    }
                    surveyState[sampleNo].photoFiles.push(file);
                    surveyState[sampleNo].photoStatus = 'O';
                    photoFileCount++;
                }
            }
        });

        // 야장파일과 사진폴더 분류가 완전히 끝났는지 체크
        if (Object.keys(surveyState).length === 0) {
            addLog('error', '분석 가능한 야장 엑셀 파일 또는 사진 폴더를 찾지 못했습니다. 폴더 구조를 다시 확인해주세요.');
            alert('야장폴더와 사진폴더가 포함된 상위 폴더 구조가 올바른지 확인해주세요!');
            return;
        }

        // 대시보드 표시 및 숫자 채우기
        dashboardSection.classList.remove('hidden');
        resultsSection.classList.remove('hidden');
        
        countExcelEl.textContent = excelCount;
        countPhotosEl.textContent = Object.values(surveyState).filter(item => item.photoFiles.length > 0).length;

        addLog('success', `분석 준비 완료: 야장 파일 ${excelCount}개, 사진 수집 폴더 ${countPhotosEl.textContent}개.`);

        // 분석 엔진 기동
        await runAnalysisEngine();
    }

    // 상태 아이템 초기화 구조
    function initSurveyItem(sampleNo) {
        surveyState[sampleNo] = {
            key: sampleNo,
            excelFile: null,
            photoFiles: [],
            excelStatus: '-',
            photoStatus: '-',
            // 추출 정보
            sampleNo: sampleNo,
            date: '',
            leader: '',
            members: '',
            pPhotoCount: 0,
            kPhotoCount: 0,
            lidar: '', // 라이다촬영 (비어있음)
            memo: '',  // 비고 (비어있음)
            hasError: false,
            warnings: []
        };
    }

    // 6. 데이터 분석 엔진 작동 (비동기 처리)
    async function runAnalysisEngine() {
        processStatusEl.textContent = '분석 진행 중...';
        processStatusEl.className = 'status-indicator processing';
        
        const keys = Object.keys(surveyState);
        const total = keys.length;
        let current = 0;
        let matchedCount = 0;
        let unmatchedCount = 0;

        for (const sampleNo of keys) {
            current++;
            const item = surveyState[sampleNo];
            
            // 진행도 계산 및 UI 갱신
            const percent = Math.round((current / total) * 100);
            progressBar.style.width = `${percent}%`;
            progressText.textContent = `진행 중... ${current} / ${total} (${percent}%)`;
            processingFileEl.textContent = `분석 중: ${sampleNo}`;

            addLog('info', `[${sampleNo}] 분석을 시작합니다.`);

            // 검증 1: 야장 엑셀과 사진 폴더 매칭 상태 확인
            if (item.excelStatus === 'O' && item.photoStatus === 'O') {
                matchedCount++;
                addLog('success', `[${sampleNo}] 야장 엑셀과 사진 폴더 매칭 성공.`);
            } else {
                unmatchedCount++;
                item.hasError = true;
                if (item.excelStatus === '-') {
                    item.warnings.push('야장 엑셀 파일 누락');
                    addLog('warning', `[${sampleNo}] 야장 엑셀 파일이 누락되었습니다. (사진 폴더만 존재)`);
                }
                if (item.photoStatus === '-') {
                    item.warnings.push('사진 폴더 누락');
                    addLog('warning', `[${sampleNo}] 사진 폴더가 누락되었습니다. (야장 파일만 존재)`);
                }
            }

            // 검증 2: 사진 카운팅 (P_ 및 K_)
            if (item.photoFiles.length > 0) {
                let pCount = 0;
                let kCount = 0;
                
                item.photoFiles.forEach(file => {
                    const name = file.name.toUpperCase();
                    if (name.startsWith('P_')) {
                        pCount++;
                    } else if (name.startsWith('K_')) {
                        kCount++;
                    }
                });

                item.pPhotoCount = pCount;
                item.kPhotoCount = kCount;
                addLog('info', `[${sampleNo}] 사진 집계: 표본점사진(P_)=${pCount}장, 구조물사진(K_)=${kCount}장`);
            }

            // 검증 3: 엑셀 파일이 존재하면 파싱하여 조사정보 추출
            if (item.excelFile) {
                try {
                    const parsedData = await parseExcelSurvey(item.excelFile);
                    
                    // 엑셀에서 추출한 데이터 맵핑
                    item.date = parsedData.date || '';
                    item.leader = parsedData.leader || '';
                    item.members = parsedData.members || '';
                    
                    // 표본점 번호 교차 검증
                    if (parsedData.sampleNo && parsedData.sampleNo !== sampleNo) {
                        item.warnings.push(`표본점번호 불일치 (파일명: ${sampleNo}, 시트내부: ${parsedData.sampleNo})`);
                        addLog('warning', `[${sampleNo}] 파일명과 시트 내의 표본점 번호가 일치하지 않습니다. (시트: ${parsedData.sampleNo})`);
                    }

                    if (!item.date) {
                        item.warnings.push('조사일자 누락');
                        addLog('warning', `[${sampleNo}] 일반·토지현황조사표 탭에서 '조사일자'를 찾을 수 없습니다.`);
                    }
                    if (!item.leader) {
                        item.warnings.push('팀장 정보 누락');
                        addLog('warning', `[${sampleNo}] 일반·토지현황조사표 탭에서 '팀장' 정보를 찾을 수 없습니다.`);
                    }
                    if (!item.members) {
                        item.warnings.push('팀원 정보 누락');
                        addLog('warning', `[${sampleNo}] 일반·토지현황조사표 탭에서 '팀원/조사자' 정보를 찾을 수 없습니다.`);
                    }

                } catch (err) {
                    item.hasError = true;
                    item.warnings.push('엑셀 파싱 실패');
                    addLog('error', `[${sampleNo}] 엑셀 파일을 읽는 동안 오류가 발생했습니다: ${err.message}`);
                }
            }

            // 브라우저 렌더링 스레드를 방해하지 않기 위해 살짝 딜레이 부여
            await new Promise(resolve => requestAnimationFrame(resolve));
        }

        // 조사일자 오름차순(과거순)으로 정렬하여 화면 일괄 렌더링
        tableBody.innerHTML = '';
        const sortedItems = Object.values(surveyState).sort((a, b) => {
            const dateA = String(a.date || '').replace(/[^0-9]/g, '');
            const dateB = String(b.date || '').replace(/[^0-9]/g, '');
            
            if (!dateA && !dateB) return 0;
            if (!dateA) return 1; // 날짜가 유실된 아이템은 뒤로 보냄
            if (!dateB) return -1;
            
            return dateA.localeCompare(dateB);
        });

        sortedItems.forEach(item => {
            renderTableRow(item);
        });

        // 분석 완료 상태 업데이트
        countMatchedEl.textContent = matchedCount;
        countUnmatchedEl.textContent = unmatchedCount;
        
        processStatusEl.textContent = '분석 완료';
        processStatusEl.className = 'status-indicator completed';
        processingFileEl.textContent = '모든 파일의 분석이 완료되었습니다.';
        addLog('success', `전체 분석 작업이 완료되었습니다. 결과표 엑셀 파일 생성이 가능합니다.`);
    }

    // 7. SheetJS를 활용한 야장 파일 내 시트 파싱
    function parseExcelSurvey(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // "일반·토지현황조사표" 시트 검색
                    const sheetName = workbook.SheetNames.find(name => 
                        name.replace(/\s+/g, '').includes('일반·토지현황조사표') || 
                        name.replace(/\s+/g, '').includes('일반토지현황')
                    );
                    
                    if (!sheetName) {
                        throw new Error('"일반·토지현황조사표" 탭을 찾을 수 없습니다.');
                    }
                    
                    const sheet = workbook.Sheets[sheetName];
                    
                    // 정보 추출: 1행 헤더, 2행 데이터 구조이므로 헤더 아래 셀의 값을 가져옴
                    const result = {
                        sampleNo: findValueBelowHeader(sheet, ['표본점번호', '표본점 번호']),
                        date: formatSurveyDate(findValueBelowHeader(sheet, ['조사일자', '조사일', '조사일시'])),
                        leader: findValueBelowHeader(sheet, ['팀장', '조사팀장', '팀 장']),
                        members: findValueBelowHeader(sheet, ['팀원', '조사원', '조사자', '조사팀원'])
                    };
                    
                    resolve(result);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = () => reject(new Error('파일 읽기 실패'));
            reader.readAsArrayBuffer(file);
        });
    }

    // 컬럼 헤더 키워드 매칭 후 그 바로 아래 행(Row + 1)의 셀 값을 가져오는 헬퍼
    function findValueBelowHeader(sheet, keywords) {
        if (!sheet['!ref']) return null;
        const range = XLSX.utils.decode_range(sheet['!ref']);
        
        for (let r = range.s.r; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cellAddress = XLSX.utils.encode_cell({ r, c });
                const cell = sheet[cellAddress];
                
                if (cell && cell.v !== undefined && cell.v !== null) {
                    const cellValStr = String(cell.v).replace(/\s+/g, '');
                    
                    // 키워드 매칭
                    const isMatched = keywords.some(keyword => cellValStr.includes(keyword));
                    
                    if (isMatched) {
                        // 헤더 매칭 성공 시 바로 아래 행(r + 1)의 셀을 가져옴
                        // 아래 방향으로 1~2칸 탐색하여 가장 먼저 만나는 비어있지 않은 값을 추출
                        for (let offset = 1; offset <= 2; offset++) {
                            if (r + offset <= range.e.r) {
                                const targetAddress = XLSX.utils.encode_cell({ r: r + offset, c });
                                const targetCell = sheet[targetAddress];
                                if (targetCell && targetCell.v !== undefined && targetCell.v !== null && String(targetCell.v).trim() !== '') {
                                    return String(targetCell.v).trim();
                                }
                            }
                        }
                    }
                }
            }
        }
        return null;
    }

    // 조사일자 YYYYMMDD 포맷 정규화
    function formatSurveyDate(dateStr) {
        if (!dateStr) return '';
        
        // 숫자만 남기기
        let clean = dateStr.replace(/[^0-9]/g, '');
        
        // 만약 엑셀 시리얼 날짜 포맷 형태인 경우 (5자리 숫자)
        if (clean.length === 5 && !isNaN(clean)) {
            const excelDate = parseInt(clean);
            // 엑셀 날짜 기준일(1899-12-30)에서 경과일 적용
            const date = new Date((excelDate - 25569) * 86400 * 1000);
            const y = date.getFullYear();
            const m = String(date.getMonth() + 1).padStart(2, '0');
            const d = String(date.getDate()).padStart(2, '0');
            return `${y}${m}${d}`;
        }

        // 일반적인 8자리 날짜 형태인 경우
        if (clean.length === 8) {
            return clean;
        }

        // 6자리 형태 (YYMMDD) 인 경우
        if (clean.length === 6) {
            const prefix = parseInt(clean.substring(0, 2)) > 50 ? '19' : '20';
            return prefix + clean;
        }

        // 그 외에는 원본 문자열에서 가능한 8자리 숫자를 파싱
        const match = dateStr.match(/\d{4}[-./\s]?\d{1,2}[-./\s]?\d{1,2}/);
        if (match) {
            const parts = match[0].split(/[-./\s]?/).filter(Boolean);
            if (parts.length >= 3) {
                const y = parts[0];
                const m = parts[1].padStart(2, '0');
                const d = parts[2].padStart(2, '0');
                return `${y}${m}${d}`;
            }
        }

        return dateStr; // 변환 불가 시 원래 값 반환
    }

    // 8. 데이터 미리보기 테이블 행 그리기
    function renderTableRow(item) {
        // 이미 렌더링된 행이 있는지 확인 후 있으면 제거 (업데이트 대응)
        const existingRow = document.getElementById(`row-${item.key}`);
        if (existingRow) {
            existingRow.remove();
        }

        const tr = document.createElement('tr');
        tr.id = `row-${item.key}`;

        // 경고 및 에러 표시 스타일
        if (item.hasError) {
            tr.className = 'row-error';
        } else if (item.warnings.length > 0) {
            tr.className = 'row-warning';
        }

        // 숫자 값은 0이면 '-'로 표시
        const pPhotoVal = item.pPhotoCount > 0 ? item.pPhotoCount : '-';
        const kPhotoVal = item.kPhotoCount > 0 ? item.kPhotoCount : '-';

        // 경고 메시지 비고란 자동 완성
        const memoContent = item.warnings.length > 0 ? `⚠️ ${item.warnings.join(', ')}` : '';

        tr.innerHTML = `
            <td><strong>${item.sampleNo}</strong></td>
            <td>${item.date || '-'}</td>
            <td>조사가능</td>
            <td>${item.leader || '-'}</td>
            <td>${item.members || '-'}</td>
            <td><span class="badge ${item.excelStatus === 'O' ? 'text-green bg-green-light' : 'text-red bg-red-light'}">${item.excelStatus}</span></td>
            <td><span class="badge ${item.photoStatus === 'O' ? 'text-green bg-green-light' : 'text-red bg-red-light'}">${item.photoStatus}</span></td>
            <td>${pPhotoVal}</td>
            <td>${kPhotoVal}</td>
            <td>${item.lidar || '-'}</td>
            <td class="text-left" style="font-size: 0.8rem; text-align: left;">${memoContent}</td>
        `;

        tableBody.appendChild(tr);
    }

    // 9. 초기화 및 리셋 단추 작동
    btnReset.addEventListener('click', () => {
        if (confirm('모든 분석 데이터를 지우고 처음으로 돌아가시겠습니까?')) {
            surveyState = {};
            logLogs = [];
            logList.innerHTML = '';
            tableBody.innerHTML = '';
            
            dashboardSection.classList.add('hidden');
            resultsSection.classList.add('hidden');
            folderInput.value = '';
            
            progressBar.style.width = '0%';
            progressText.textContent = '준비 중... (0%)';
            processingFileEl.textContent = '';
            
            addLog('info', '화면을 초기화했습니다. 폴더를 새로 드래그 해주세요.');
        }
    });



    // 10. 통합 결과표 엑셀 파일 생성 및 다운로드 (핵심)
    btnDownload.addEventListener('click', () => {
        try {
            const rawItems = Object.values(surveyState);
            if (rawItems.length === 0) {
                alert('다운로드할 데이터가 없습니다. 폴더를 먼저 분석해주세요.');
                return;
            }

            // 조사일자 오름차순(과거순) 정렬 처리
            const items = [...rawItems].sort((a, b) => {
                const dateA = String(a.date || '').replace(/[^0-9]/g, '');
                const dateB = String(b.date || '').replace(/[^0-9]/g, '');
                
                if (!dateA && !dateB) return 0;
                if (!dateA) return 1;
                if (!dateB) return -1;
                
                return dateA.localeCompare(dateB);
            });

            // 인앱 웹뷰/개발도구 미리보기 샌드박스로 인해 a[download] 속성이 무력화되는 경우를 방지하기 위해 경고 및 가이드 제공
            const isProbablySandbox = window.self !== window.top || 
                                     /ProjectIDX|Gemini/i.test(navigator.userAgent) ||
                                     window.name === 'preview' ||
                                     (window.location.ancestorOrigins && window.location.ancestorOrigins.length > 0);
            
            if (isProbablySandbox) {
                alert(`💡 [안내] 엑셀 파일 다운로드 방법\n\n현재 'Gemini/IDX 인앱 브라우저' 뷰어 환경에서 다운로드 버튼을 클릭하셨습니다.\n\n이 뷰어 환경은 보안 정책상 다운로드 파일명 지정을 차단하므로, 다운로드 폴더에 확장자(.xlsx)가 없는 임시 난수명(UUID)으로 저장될 수 있습니다.\n\n가장 깔끔하고 정상적인 다운로드를 위해 아래 방법을 이용해 주세요:\n\n1️⃣ 크롬(Chrome) 브라우저를 열고 주소창에 직접 입력 접속: http://localhost:3000\n2️⃣ 만약 현재 뷰어에서 그대로 받으셨다면, 다운로드 폴더에서 받아진 임시 파일 이름 뒤에 수동으로 '.xlsx'를 덧붙여주시면 엑셀 파일로 바로 열립니다!`);
            }

            addLog('info', '통합 결과표 엑셀 파일 생성을 시작합니다.');

            // 엑셀을 위한 2차원 배열 데이터 생성 (라이더촬영, 비고 열 제외 - 총 8개 열)
            const headers = [
                '표본점번호', // A (Index 0)
                '조사일자',   // B (Index 1)
                '조사가능불가', // C (Index 2)
                '팀장',       // D (Index 3)
                '팀원',       // E (Index 4)
                '엑셀',       // F (Index 5)
                '사진',       // G (Index 6)
                '표본점사진', // H (Index 7)
                '구조물사진'  // I (Index 8)
            ];

            const aoaData = [
                headers // 첫번째 행에 헤더 추가
            ];

            // 데이터 행 추가
            items.forEach(item => {
                // 표본점번호를 숫자로 변환 (변환 불가능한 경우 대비하여 안전 장치 적용)
                const sampleNoNum = Number(item.sampleNo);
                const sampleNoVal = isNaN(sampleNoNum) ? item.sampleNo : sampleNoNum;

                const row = [
                    sampleNoVal,
                    item.date || '',
                    '조사가능', // 조사가능불가
                    item.leader || '',
                    item.members || '',
                    item.excelStatus,
                    item.photoStatus,
                    item.pPhotoCount > 0 ? item.pPhotoCount : '-',
                    item.kPhotoCount > 0 ? item.kPhotoCount : '-'
                ];
                
                aoaData.push(row);
            });

            // 워크북 생성
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(aoaData);

            // 모든 데이터 셀에 가운데 정렬 스타일 적용, 테두리 추가 및 표본점번호 서식 지정
            for (const key in ws) {
                if (key[0] !== '!') { // !ref, !cols 등 메타 데이터 제외
                    ws[key].s = {
                        alignment: {
                            horizontal: 'center',
                            vertical: 'center'
                        },
                        border: {
                            top: { style: 'thin', color: { rgb: '000000' } },
                            bottom: { style: 'thin', color: { rgb: '000000' } },
                            left: { style: 'thin', color: { rgb: '000000' } },
                            right: { style: 'thin', color: { rgb: '000000' } }
                        }
                    };

                    // A열(표본점번호)이고 A1(헤더)이 아닌 경우 숫자 서식(z: '0') 적용
                    if (key.startsWith('A') && key !== 'A1') {
                        ws[key].z = '0';
                    }
                }
            }

            // 시트 서식 및 레이아웃 튜닝 (열 너비 조절 - 9개 열 규격)
            const colWidths = [
                { wch: 15 }, // A: 표본점번호
                { wch: 12 }, // B: 조사일자
                { wch: 15 }, // C: 조사가능불가
                { wch: 10 }, // D: 팀장
                { wch: 15 }, // E: 팀원
                { wch: 8 },  // F: 엑셀
                { wch: 8 },  // G: 사진
                { wch: 12 }, // H: 표본점사진
                { wch: 12 }  // I: 구조물사진
            ];
            ws['!cols'] = colWidths;

            XLSX.utils.book_append_sheet(wb, ws, '조사결과표');

            // 최상위 폴더명 타입 체크 및 UUID 우회 안전 조치
            const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
            let checkFolderName = '';
            
            if (typeof topFolderName === 'string') {
                checkFolderName = topFolderName.trim();
            }
            
            let safeFolderName = '도시산림자원조사';
            if (checkFolderName !== '' && !uuidRegex.test(checkFolderName) && checkFolderName !== 'null' && checkFolderName !== 'undefined') {
                safeFolderName = checkFolderName;
            }

            const fileName = `${safeFolderName}_도시산림자원조사 결과표.xlsx`;

            // 최적화된 다운로드 링크 수동 생성 (Blob 처리)
            // XLSX.writeFile 대신 이 브라우저 친화적 엔진이 다운로드 샌드박스를 완벽 우회합니다.
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
            
            function s2ab(s) {
                const buf = new ArrayBuffer(s.length);
                const view = new Uint8Array(buf);
                for (let i = 0; i < s.length; i++) {
                    view[i] = s.charCodeAt(i) & 0xFF;
                }
                return buf;
            }

            const blob = new Blob([s2ab(wbout)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const blobUrl = URL.createObjectURL(blob);
            
            const downloadLink = document.createElement('a');
            downloadLink.href = blobUrl;
            downloadLink.download = fileName;
            
            document.body.appendChild(downloadLink);
            downloadLink.click();
            
            // 뒷정리
            document.body.removeChild(downloadLink);
            setTimeout(() => {
                URL.revokeObjectURL(blobUrl);
            }, 100);
            
            addLog('success', `통합 결과표 파일 다운로드를 트리거했습니다. 파일명: ${fileName}`);
        } catch (error) {
            console.error(error);
            addLog('error', `엑셀 파일 생성 중 치명적 오류가 발생했습니다: ${error.message}`);
            alert(`다운로드 실패: ${error.message}\n화면 하단의 실시간 로그 콘솔을 확인해주세요.`);
        }
    });

    // 11. 브라우저 화면 빈 공간에 폴더/파일 드롭 시 페이지가 이탈하는 현상 방지
    window.addEventListener('dragover', (e) => {
        e.preventDefault();
    }, false);
    window.addEventListener('drop', (e) => {
        e.preventDefault();
    }, false);
});
