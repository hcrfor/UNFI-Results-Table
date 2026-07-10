const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// 1. 테스트 기본 경로 설정
const baseDir = path.join(__dirname, 'test_dataset');
const excelDir = path.join(baseDir, '야장폴더');
const photoDir = path.join(baseDir, '사진폴더');

// 디렉토리 자동 생성 함수
function ensureDirectoryExistence(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
    }
}

ensureDirectoryExistence(excelDir);
ensureDirectoryExistence(photoDir);

console.log('테스트 데이터셋 구조 생성 시작...');

// 2. 가상의 표본점 데이터 리스트
const mockPlots = [
    { sampleNo: '410977119997', date: '2026-07-07', leader: '한성안', member: '한승훈', pPhotos: 5, kPhotos: 3 },
    { sampleNo: '410977120007', date: '2026.07.07', leader: '한성안', member: '한승훈', pPhotos: 5, kPhotos: 3 },
    { sampleNo: '410977320001', date: '2026년 7월 7일', leader: '한성안', member: '한승훈', pPhotos: 5, kPhotos: 3 },
    { sampleNo: '410977820012', date: '2026/07/07', leader: '한성안', member: '한승훈', pPhotos: 5, kPhotos: 3 },
    { sampleNo: '410977820014', date: '20260707', leader: '한성안', member: '한승훈', pPhotos: 5, kPhotos: 3 },
    { sampleNo: '410978220174', date: '20260707', leader: '한성안', member: '한승훈', pPhotos: 5, kPhotos: 0 } // K_사진이 없는 상태 테스트
];

// 3. 야장 엑셀 파일 및 사진 파일들 생성 루프
mockPlots.forEach(plot => {
    // A. 엑셀 파일 경로
    const excelFilePath = path.join(excelDir, `${plot.sampleNo}.xlsx`);
    
    // 워크북 객체 빌드
    const wb = XLSX.utils.book_new();
    
    // "일반·토지현황조사표" 시트에 들어갈 모의 셀 배열 데이터
    // 1행 헤더, 2행 데이터 구조로 생성합니다.
    const rowData = [
        ['표본점번호', '조사일자', '소속', '팀장', '팀원'],
        [plot.sampleNo, plot.date, '산림자원조사본부', plot.leader, plot.member]
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(rowData);
    XLSX.utils.book_append_sheet(wb, ws, '일반·토지현황조사표');
    XLSX.writeFile(wb, excelFilePath);
    console.log(`[야장 생성 완료] ${excelFilePath}`);

    // B. 사진 폴더 경로 및 가상 이미지 파일 생성
    const plotPhotoDir = path.join(photoDir, plot.sampleNo);
    ensureDirectoryExistence(plotPhotoDir);

    // P_로 시작하는 표본점 사진 생성
    for (let i = 1; i <= plot.pPhotos; i++) {
        const photoPath = path.join(plotPhotoDir, `P_IMG_${plot.date.replace(/[^0-9]/g, '')}_00${i}.jpg`);
        fs.writeFileSync(photoPath, 'fake_image_data_P');
    }

    // K_로 시작하는 구조물 사진 생성
    for (let i = 1; i <= plot.kPhotos; i++) {
        const photoPath = path.join(plotPhotoDir, `K_IMG_${plot.date.replace(/[^0-9]/g, '')}_00${i}.jpg`);
        fs.writeFileSync(photoPath, 'fake_image_data_K');
    }
    
    console.log(`[사진 폴더 생성 완료] ${plotPhotoDir} (P:${plot.pPhotos}장, K:${plot.kPhotos}장)`);
});

// 4. 매칭 오류 테스트를 위한 예외 테스트 케이스 추가
// 야장은 있고 사진폴더가 없는 케이스
const errPlot1 = '410979119999';
const errExcelFilePath = path.join(excelDir, `${errPlot1}.xlsx`);
const wbErr = XLSX.utils.book_new();
const wsErr = XLSX.utils.aoa_to_sheet([
    ['표본점번호', '조사일자', '소속', '팀장', '팀원'],
    [errPlot1, '20260708', '산림자원조사본부', '김선생', '이선생']
]);
XLSX.utils.book_append_sheet(wbErr, wsErr, '일반·토지현황조사표');
XLSX.writeFile(wbErr, errExcelFilePath);
console.log(`[오류 테스트용 야장 생성] ${errExcelFilePath} (사진 폴더 없음)`);

// 사진폴더는 있고 야장이 없는 케이스
const errPlot2 = '410979220000';
const errPhotoDir = path.join(photoDir, errPlot2);
ensureDirectoryExistence(errPhotoDir);
fs.writeFileSync(path.join(errPhotoDir, 'P_001.jpg'), 'fake');
fs.writeFileSync(path.join(errPhotoDir, 'K_001.jpg'), 'fake');
console.log(`[오류 테스트용 사진 폴더 생성] ${errPhotoDir} (야장 엑셀 파일 없음)`);

console.log('\n모든 가상 테스트 데이터셋이 "test_dataset" 폴더 내에 성공적으로 빌드되었습니다.');
console.log('이 폴더를 웹 브라우저 화면의 업로드 영역으로 드래그하여 바로 실시간 테스트할 수 있습니다.');
