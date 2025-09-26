class AttendanceCalculator {
    constructor() {
        this.initEventListeners();
        this.attendanceData = [];
        this.debugLog = [];
        this.holidays = new Set();
        this.loadHolidays();
    }

    initEventListeners() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const toggleDebug = document.getElementById('toggleDebug');

        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', this.handleDragOver.bind(this));
        uploadArea.addEventListener('dragleave', this.handleDragLeave.bind(this));
        uploadArea.addEventListener('drop', this.handleDrop.bind(this));
        fileInput.addEventListener('change', this.handleFileSelect.bind(this));

        if (toggleDebug) {
            toggleDebug.addEventListener('click', this.toggleDebugInfo.bind(this));
        }
    }

    addDebugLog(message, data = null) {
        const timestamp = new Date().toLocaleTimeString();
        const logEntry = { timestamp, message, data };
        this.debugLog.push(logEntry);
        console.log(`[${timestamp}] ${message}`, data || '');
    }

    toggleDebugInfo() {
        const debugInfo = document.getElementById('debugInfo');
        const toggleButton = document.getElementById('toggleDebug');

        if (debugInfo.style.display === 'none') {
            debugInfo.style.display = 'block';
            toggleButton.textContent = '디버깅 정보 숨기기';
            this.updateDebugInfo();
        } else {
            debugInfo.style.display = 'none';
            toggleButton.textContent = '디버깅 정보 표시';
        }
    }

    updateDebugInfo() {
        const debugInfo = document.getElementById('debugInfo');
        const logText = this.debugLog.map(log => {
            let text = `[${log.timestamp}] ${log.message}`;
            if (log.data) {
                text += '\n' + JSON.stringify(log.data, null, 2);
            }
            return text;
        }).join('\n\n');

        debugInfo.textContent = logText;
        debugInfo.scrollTop = debugInfo.scrollHeight;
    }

    async loadHolidays() {
        try {
            // 한국천문연구원 특일정보 API 사용
            const year = new Date().getFullYear();
            const response = await fetch(`https://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo?serviceKey=YOUR_API_KEY&solYear=${year}&_type=json`);

            if (!response.ok) {
                throw new Error('Holiday API failed');
            }

            const data = await response.json();
            if (data.response && data.response.body && data.response.body.items) {
                data.response.body.items.item.forEach(holiday => {
                    const dateStr = holiday.locdate.toString();
                    const year = parseInt(dateStr.substring(0, 4));
                    const month = parseInt(dateStr.substring(4, 6)) - 1;
                    const day = parseInt(dateStr.substring(6, 8));
                    const holidayDate = new Date(year, month, day);
                    this.holidays.add(holidayDate.toDateString());
                });
            }

            this.addDebugLog('공휴일 데이터 로드 완료', { count: this.holidays.size });
        } catch (error) {
            console.warn('공휴일 API 로드 실패, 기본 공휴일 사용:', error);
            // API 실패 시 기본 공휴일 설정
            this.setDefaultHolidays();
        }
    }

    setDefaultHolidays() {
        // 2025년 기본 공휴일 설정
        const defaultHolidays = [
            '2025-01-01', // 신정
            '2025-02-09', '2025-02-10', '2025-02-11', '2025-02-12', // 설날 연휴
            '2025-03-01', // 삼일절
            '2025-05-05', // 어린이날
            '2025-05-15', // 부처님오신날
            '2025-06-06', // 현충일
            '2025-08-15', // 광복절
            '2025-10-03', '2025-10-04', '2025-10-05', '2025-10-06', // 추석 연휴
            '2025-10-09', // 한글날
            '2025-12-25' // 성탄절
        ];

        defaultHolidays.forEach(dateStr => {
            const date = new Date(dateStr);
            this.holidays.add(date.toDateString());
        });

        this.addDebugLog('기본 공휴일 설정 완료', { count: this.holidays.size });
    }


    isDefaultHoliday(date) {
        const defaultHolidays = [
            '2025-01-01', '2025-02-09', '2025-02-10', '2025-02-11', '2025-02-12',
            '2025-03-01', '2025-05-05', '2025-05-15', '2025-06-06', '2025-08-15',
            '2025-10-03', '2025-10-04', '2025-10-05', '2025-10-06',
            '2025-10-09', '2025-12-25'
        ];
        const dateString = date.toISOString().split('T')[0];
        return defaultHolidays.includes(dateString);
    }

    handleDragOver(e) {
        e.preventDefault();
        e.currentTarget.classList.add('dragover');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('dragover');
    }

    handleDrop(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const files = e.target.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    async processFile(file) {
        if (!file) return;

        const validTypes = ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
        if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx?|xls)$/i)) {
            alert('엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.');
            return;
        }

        try {
            this.debugLog = []; // 디버그 로그 초기화
            this.addDebugLog('파일 처리 시작', { name: file.name, type: file.type, size: file.size });

            const arrayBuffer = await file.arrayBuffer();
            this.addDebugLog('ArrayBuffer 생성 완료', { size: arrayBuffer.byteLength });

            const workbook = XLSX.read(arrayBuffer, {
                type: 'array',
                cellText: false,
                cellDates: true
            });

            this.addDebugLog('워크북 로드 완료', { sheets: workbook.SheetNames });

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            this.addDebugLog('워크시트 선택', { sheet: firstSheetName, range: worksheet['!ref'] });

            // 여러 형식으로 데이터 파싱 시도
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                raw: false,
                blankrows: false,
                defval: ''
            });

            this.addDebugLog('JSON 데이터 파싱 완료', { totalRows: jsonData.length, sample: jsonData.slice(0, 3) });

            // 빈 행 제거
            const filteredData = jsonData.filter(row =>
                row && row.some(cell => cell !== null && cell !== undefined && cell !== '')
            );

            this.addDebugLog('빈 행 제거 완료', { filteredRows: filteredData.length });

            if (filteredData.length < 2) {
                this.addDebugLog('데이터 부족 오류', { rows: filteredData.length });
                alert('데이터가 충분하지 않습니다. 헤더와 최소 1행의 데이터가 필요합니다.');
                document.getElementById('debugSection').style.display = 'block';
                return;
            }

            this.parseAttendanceData(filteredData);

            if (this.attendanceData.length === 0) {
                this.addDebugLog('근태 데이터 파싱 실패', { parsedData: this.attendanceData });
                alert('근태 데이터를 찾을 수 없습니다. 파일 형식을 확인해주세요.');
                document.getElementById('debugSection').style.display = 'block';
                return;
            }

            this.calculateAttendance();
            this.displayResults();
            document.getElementById('debugSection').style.display = 'block';

        } catch (error) {
            console.error('파일 처리 중 오류:', error);
            alert(`파일을 읽는 중 오류가 발생했습니다: ${error.message}`);
        }
    }

    parseAttendanceData(data) {
        this.attendanceData = [];

        if (data.length < 2) return;

        this.addDebugLog('헤더 행 분석', data[0]);

        const headers = data[0];
        let dateCol = -1, startCol = -1, endCol = -1, noteCol = -1, statusCol = -1, annualLeaveCol = -1;

        // 헤더 매칭을 더 유연하게 개선
        headers.forEach((header, index) => {
            if (header) {
                const h = header.toString().toLowerCase().trim();
                this.addDebugLog(`컬럼 ${index} 분석`, { original: header, normalized: h });

                // 날짜 컬럼 찾기
                if (h.includes('날짜') || h.includes('date') || h.includes('일') || h.includes('day') ||
                    h.match(/^\d{4}/) || h.includes('월') || h.includes('년')) {
                    dateCol = index;
                    this.addDebugLog('날짜 컬럼 발견', index);
                }
                // 출근시간 컬럼 찾기
                else if (h.includes('출근') || h.includes('시작') || h.includes('start') ||
                         h.includes('in') || h.includes('체크인') || h.includes('근무시작')) {
                    startCol = index;
                    this.addDebugLog('출근시간 컬럼 발견', index);
                }
                // 퇴근시간 컬럼 찾기
                else if (h.includes('퇴근') || h.includes('종료') || h.includes('end') ||
                         h.includes('out') || h.includes('체크아웃') || h.includes('근무종료')) {
                    endCol = index;
                    this.addDebugLog('퇴근시간 컬럼 발견', index);
                }
                // 상태 컬럼 찾기 (연차 시간이 기록될 수 있는 컬럼)
                else if (h.includes('상태') || h.includes('status') || h.includes('연차')) {
                    statusCol = index;
                    this.addDebugLog('상태 컬럼 발견', index);
                }
                // 비고 컬럼 찾기
                else if (h.includes('비고') || h.includes('메모') || h.includes('note') ||
                         h.includes('remark') || h.includes('comment')) {
                    noteCol = index;
                    this.addDebugLog('비고 컬럼 발견', index);
                }
            }
        });

        // H열 (7번 인덱스)를 연차 컬럼으로 지정
        annualLeaveCol = 7; // H열은 0부터 시작하므로 7번 인덱스
        this.addDebugLog('H열을 연차 컬럼으로 지정', annualLeaveCol);

        // 컬럼을 찾지 못한 경우 순서대로 추정
        if (dateCol === -1 && headers.length > 0) {
            dateCol = 0;
            this.addDebugLog('날짜 컬럼을 첫 번째로 추정', dateCol);
        }
        if (startCol === -1 && headers.length > 1) {
            startCol = 1;
            this.addDebugLog('출근시간 컬럼을 두 번째로 추정', startCol);
        }
        if (endCol === -1 && headers.length > 2) {
            endCol = 2;
            this.addDebugLog('퇴근시간 컬럼을 세 번째로 추정', endCol);
        }
        if (noteCol === -1 && headers.length > 3) {
            noteCol = 3;
            this.addDebugLog('비고 컬럼을 네 번째로 추정', noteCol);
        }
        if (statusCol === -1 && headers.length > 4) {
            statusCol = 4;
            this.addDebugLog('상태 컬럼을 다섯 번째로 추정', statusCol);
        }

        this.addDebugLog('최종 컬럼 매핑', { dateCol, startCol, endCol, noteCol, statusCol, annualLeaveCol });

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;

            this.addDebugLog(`행 ${i} 데이터`, row);

            const record = {
                date: this.parseDate(row[dateCol] || ''),
                startTime: this.parseTime(row[startCol] || ''),
                endTime: this.parseTime(row[endCol] || ''),
                note: (row[noteCol] || '').toString().trim(),
                status: (row[statusCol] || '').toString().trim(),
                annualLeaveHours: 0, // 초기값은 0, 나중에 감지되면 설정
                hColumnData: (row[annualLeaveCol] || '').toString().trim() // H열 원본 데이터 보존
            };

            this.addDebugLog(`행 ${i} 파싱 결과`, record);

            if (record.date) {
                this.attendanceData.push(record);
            }
        }

        // 해당 월 데이터만 필터링
        if (this.attendanceData.length > 0) {
            const firstDate = this.attendanceData[0].date;
            const targetMonth = firstDate.getMonth();
            const targetYear = firstDate.getFullYear();

            this.attendanceData = this.attendanceData.filter(record => {
                return record.date.getMonth() === targetMonth &&
                       record.date.getFullYear() === targetYear;
            });

            this.addDebugLog('해당 월 필터링 완료', {
                targetMonth: targetMonth + 1,
                targetYear: targetYear,
                filteredCount: this.attendanceData.length
            });
        }

        this.addDebugLog('최종 파싱 완료', { count: this.attendanceData.length, data: this.attendanceData });
    }

    parseDate(dateValue) {
        if (!dateValue) return null;

        console.log('날짜 파싱 시도:', dateValue, typeof dateValue);

        // Date 객체인 경우
        if (dateValue instanceof Date) {
            return dateValue;
        }

        let dateStr = dateValue.toString().trim();

        // 엑셀 시리얼 날짜 (5자리 숫자)
        if (dateStr.match(/^\d{5}$/)) {
            const excelDate = parseInt(dateStr);
            const date = new Date((excelDate - 25569) * 86400 * 1000);
            console.log('엑셀 시리얼 날짜 변환:', dateStr, '->', date);
            return date;
        }

        // 다양한 날짜 형식 지원
        const patterns = [
            /^(\d{4})-(\d{1,2})-(\d{1,2})$/,          // 2025-09-01
            /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/,        // 2025/09/01
            /^(\d{4})\.(\d{1,2})\.(\d{1,2})$/,        // 2025.09.01
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,        // 09/01/2025
            /^(\d{1,2})-(\d{1,2})-(\d{4})$/,          // 09-01-2025
            /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/,        // 09.01.2025
            /^(\d{4})(\d{2})(\d{2})$/,                // 20250901
            /^(\d{2})(\d{2})(\d{2})$/                 // 250901
        ];

        for (let pattern of patterns) {
            const match = dateStr.match(pattern);
            if (match) {
                let year, month, day;

                if (pattern.source.includes('(\\d{4})') && pattern.source.startsWith('^(\\d{4})')) {
                    // YYYY-MM-DD 형식
                    [, year, month, day] = match;
                } else if (pattern.source.includes('(\\d{4})') && pattern.source.endsWith('(\\d{4})$')) {
                    // MM-DD-YYYY 형식
                    [, month, day, year] = match;
                } else if (pattern.source === '^(\\d{4})(\\d{2})(\\d{2})$') {
                    // YYYYMMDD 형식
                    [, year, month, day] = match;
                } else if (pattern.source === '^(\\d{2})(\\d{2})(\\d{2})$') {
                    // YYMMDD 형식
                    [, year, month, day] = match;
                    year = '20' + year; // 20을 앞에 붙여서 2025년으로 변환
                }

                const parsedDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                console.log('패턴 매칭 날짜 변환:', dateStr, '->', parsedDate);
                return parsedDate;
            }
        }

        // 기본 Date 생성자 시도
        const date = new Date(dateStr);
        if (!isNaN(date.getTime())) {
            console.log('기본 Date 생성자로 날짜 변환:', dateStr, '->', date);
            return date;
        }

        console.log('날짜 파싱 실패:', dateStr);
        return null;
    }

    parseTime(timeValue) {
        if (!timeValue) return null;

        console.log('시간 파싱 시도:', timeValue, typeof timeValue);

        // Date 객체인 경우 시간 추출
        if (timeValue instanceof Date) {
            return {
                hours: timeValue.getHours(),
                minutes: timeValue.getMinutes()
            };
        }

        let timeStr = timeValue.toString().trim();

        // ISO 날짜 문자열 (예: 2025-09-01T09:00:00.000Z)
        if (timeStr.includes('T')) {
            const date = new Date(timeStr);
            if (!isNaN(date.getTime())) {
                console.log('ISO 날짜에서 시간 추출:', timeStr, '->', date.getHours(), ':', date.getMinutes());
                return {
                    hours: date.getHours(),
                    minutes: date.getMinutes()
                };
            }
        }

        // 일반적인 시간 형식 (HH:MM 또는 HH:MM:SS)
        const timePattern = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/;
        const match = timeStr.match(timePattern);

        if (match) {
            console.log('시간 패턴 매칭:', timeStr, '->', match[1], ':', match[2]);
            return {
                hours: parseInt(match[1]),
                minutes: parseInt(match[2])
            };
        }

        // 엑셀 시간 소수점 형식 (0.375 = 9시간)
        const decimalMatch = timeStr.match(/^0?\.\d+$|^\d+\.\d+$/);
        if (decimalMatch) {
            const decimal = parseFloat(timeStr);
            const hours = Math.floor(decimal * 24);
            const minutes = Math.round((decimal * 24 - hours) * 60);
            console.log('소수점 시간 변환:', timeStr, '->', hours, ':', minutes);
            return { hours, minutes };
        }

        // 4자리 숫자 형식 (0900, 1730 등)
        const fourDigitMatch = timeStr.match(/^(\d{2})(\d{2})$/);
        if (fourDigitMatch) {
            const hours = parseInt(fourDigitMatch[1]);
            const minutes = parseInt(fourDigitMatch[2]);
            if (hours < 24 && minutes < 60) {
                console.log('4자리 시간 변환:', timeStr, '->', hours, ':', minutes);
                return { hours, minutes };
            }
        }

        // 3자리 숫자 형식 (900, 1730 등)
        const threeDigitMatch = timeStr.match(/^(\d{1})(\d{2})$/);
        if (threeDigitMatch) {
            const hours = parseInt(threeDigitMatch[1]);
            const minutes = parseInt(threeDigitMatch[2]);
            if (hours < 24 && minutes < 60) {
                console.log('3자리 시간 변환:', timeStr, '->', hours, ':', minutes);
                return { hours, minutes };
            }
        }

        console.log('시간 파싱 실패:', timeStr);
        return null;
    }

    calculateWorkHours(startTime, endTime) {
        if (!startTime || !endTime) return 0;

        const startMinutes = startTime.hours * 60 + startTime.minutes;
        const endMinutes = endTime.hours * 60 + endTime.minutes;

        let workMinutes = endMinutes - startMinutes;
        if (workMinutes < 0) {
            workMinutes += 24 * 60;
        }

        return workMinutes / 60;
    }

    calculateBreakTime(workHours) {
        // L15 수식 참고: =IF(G9=0,,IF(HOUR($G9)>=4.5,1,IF(HOUR($G9)<4.5,0.5,)))
        if (workHours === 0) {
            return 0;
        }

        if (workHours >= 4.5) {
            return 1; // 4.5시간 이상이면 1시간 휴게
        } else if (workHours < 4.5) {
            return 0.5; // 4.5시간 미만이면 0.5시간 휴게
        }

        return 0;
    }

    isWeekend(date) {
        const day = date.getDay();
        return day === 0 || day === 6;
    }

    isHoliday(date) {
        return this.holidays.has(date.toDateString());
    }

    calculateAttendance() {
        let totalWorkHours = 0;
        let totalBreakHours = 0;
        let totalRegularHours = 0;
        let totalOvertimeHours = 0;
        let annualLeaveDays = 0;
        let workDays = 0;
        let totalActualWorkHours = 0; // 실근무시간 합계를 별도로 계산

        const processedData = this.attendanceData.map(record => {
            const result = {
                ...record,
                workHours: 0,
                breakHours: 0,
                actualWorkHours: 0,
                status: '정상'
            };

            const note = record.note.toLowerCase();
            const status = record.status ? record.status.toLowerCase() : '';

            this.addDebugLog(`${this.formatDate(record.date)} 연차 검사`, {
                originalNote: record.note,
                originalStatus: record.status,
                annualLeaveHours: record.annualLeaveHours,
                lowerNote: note,
                lowerStatus: status
            });

            // H열에서 연차 시간 확인 (최우선) - "완료 ( 연차 8.00h )" 형식
            if (record.hColumnData && record.hColumnData.trim() !== '') {
                const hColumnValue = record.hColumnData.trim();

                this.addDebugLog(`${this.formatDate(record.date)} H열 데이터`, {
                    rawValue: hColumnValue
                });

                // "완료 ( 연차 8.00h )" 또는 "연차 8.00h" 형식에서 숫자 추출
                const annualLeavePatterns = [
                    /완료\s*\(\s*연차\s*(\d+(?:\.\d+)?)h?\s*\)/i, // 완료 ( 연차 8.00h )
                    /완료\s*\(\s*연차\s*(\d+(?:\.\d+)?)\s*\)/i,   // 완료 ( 연차 8 )
                    /연차\s*(\d+(?:\.\d+)?)h?/i,                  // 연차 8h 또는 연차 8
                    /(\d+(?:\.\d+)?)h?\s*연차/i                   // 8h 연차 또는 8 연차
                ];

                let foundAnnualLeave = false;
                for (let pattern of annualLeavePatterns) {
                    const match = hColumnValue.match(pattern);
                    if (match) {
                        const hours = parseFloat(match[1]);
                        const days = hours / 8;

                        result.annualLeaveHours = hours; // 연차시간 별도 저장
                        result.status = '연차';
                        annualLeaveDays += days;
                        foundAnnualLeave = true;

                        this.addDebugLog(`${this.formatDate(record.date)} H열 연차 감지`, {
                            pattern: pattern.source,
                            rawValue: hColumnValue,
                            matchedValue: match[1],
                            hours: hours,
                            days: days,
                            totalAnnualLeave: annualLeaveDays
                        });

                        break;
                    }
                }

                // 단순히 숫자만 있는 경우도 처리 (8.00, 4.00 등)
                if (!foundAnnualLeave) {
                    const simpleNumberMatch = hColumnValue.match(/^(\d+(?:\.\d+)?)$/);
                    if (simpleNumberMatch) {
                        const hours = parseFloat(simpleNumberMatch[1]);
                        if (hours > 0) {
                            const days = hours / 8;

                            result.annualLeaveHours = hours;
                            result.status = '연차';
                            annualLeaveDays += days;
                            foundAnnualLeave = true;

                            this.addDebugLog(`${this.formatDate(record.date)} H열 단순 숫자 연차 감지`, {
                                rawValue: hColumnValue,
                                hours: hours,
                                days: days,
                                totalAnnualLeave: annualLeaveDays
                            });
                        }
                    }
                }

                // H열에서 연차가 감지되었다면 다른 연차 검사는 스킵
                if (foundAnnualLeave) {
                    // 여기서는 return하지 않고 계속 진행하여 기본 근무시간과 합산되도록 함
                } else {
                    // H열에서 연차를 찾지 못한 경우에만 다른 검사 진행

                    // 상태 컬럼에서 연차 시간 확인 (두 번째 우선순위)
                    // '완료 ( 연차 8.00h )' 형식 처리
                    const statusLeaveMatch = status.match(/연차\s*(\d+(?:\.\d+)?)/);
                    const statusCompleteMatch = status.match(/완료\s*\(\s*연차\s*(\d+(?:\.\d+)?)h?\s*\)/);

                    this.addDebugLog(`${this.formatDate(record.date)} 정규식 결과`, {
                        statusLeaveMatch: statusLeaveMatch,
                        statusCompleteMatch: statusCompleteMatch,
                        statusIncludes연차: status.includes('연차')
                    });

                    if (statusLeaveMatch || statusCompleteMatch || status.includes('연차')) {
                        result.status = '연차';
                        let hours = 8; // 기본값
                        if (statusCompleteMatch) {
                            hours = parseFloat(statusCompleteMatch[1]);
                        } else if (statusLeaveMatch) {
                            hours = parseFloat(statusLeaveMatch[1]);
                        }

                        result.annualLeaveHours = hours; // 연차시간 별도 저장
                        const days = hours / 8;
                        annualLeaveDays += days;

                        this.addDebugLog(`${this.formatDate(record.date)} 상태열 연차 감지`, {
                            hours: hours,
                            days: days,
                            totalAnnualLeave: annualLeaveDays
                        });

                        // 연차 감지 후에도 계속 진행하여 기본 근무시간과 합산하지 않고 return (상태열은 독립적으로 처리)
                        return result;
                    }

                    // 비고에서 연차 관련 처리
                    if (note.includes('연차') || note.includes('휴가') || note.includes('annual')) {
                        result.status = '연차';

                        // 숫자가 포함된 연차 시간 처리 (예: "연차 8", "연차8시간" 등)
                        const annualLeaveMatch = note.match(/연차\s*(\d+(?:\.\d+)?)/);
                        let hours = 8; // 기본값
                        if (annualLeaveMatch) {
                            hours = parseFloat(annualLeaveMatch[1]);
                            annualLeaveDays += hours / 8; // 8시간 = 1일 연차로 환산
                        } else {
                            annualLeaveDays++; // 기본 1일 연차
                        }

                        result.annualLeaveHours = hours; // 연차시간 별도 저장

                        // 비고에서 연차 감지 시에도 return (독립적으로 처리)
                        return result;
                    }
                }
            }

            // 기타 상태 처리
            if (note.includes('병가') || note.includes('반차') || note.includes('외근')) {
                result.status = note.includes('병가') ? '병가' : note.includes('반차') ? '반차' : '외근';
                if (note.includes('반차')) {
                    annualLeaveDays += 0.5;
                    result.annualLeaveHours = 4; // 반차는 4시간으로 처리
                }
                return result;
            }

            if (this.isWeekend(record.date) || this.isHoliday(record.date)) {
                if (record.startTime && record.endTime) {
                    result.status = '휴일근무';
                } else {
                    result.status = '휴무';
                    return result;
                }
            }

            if (record.startTime && record.endTime) {
                const workHours = this.calculateWorkHours(record.startTime, record.endTime);
                const breakHours = this.calculateBreakTime(workHours);
                const baseActualWorkHours = Math.max(0, workHours - breakHours);
                const annualLeaveHours = result.annualLeaveHours || 0;

                result.workHours = workHours;
                result.breakHours = breakHours;
                result.actualWorkHours = baseActualWorkHours + annualLeaveHours; // 기본근무시간 + 연차시간

                totalWorkHours += workHours;
                totalBreakHours += breakHours;

                if (this.isWeekend(record.date) || this.isHoliday(record.date)) {
                    totalOvertimeHours += result.actualWorkHours;
                    result.status = '휴일근무';
                } else {
                    const regularHours = Math.min(result.actualWorkHours, 8);
                    const overtimeHours = Math.max(0, result.actualWorkHours - 8);

                    totalRegularHours += regularHours;
                    totalOvertimeHours += overtimeHours;
                    workDays++;
                    result.status = '정상';
                }
            } else if (record.startTime && !record.endTime) {
                // 출근시간만 있고 퇴근시간이 없는 경우: 9시간 후를 퇴근시간으로 자동 계산
                const autoEndTime = {
                    hours: record.startTime.hours + 9,
                    minutes: record.startTime.minutes
                };

                // 24시를 넘어가는 경우 처리
                if (autoEndTime.hours >= 24) {
                    autoEndTime.hours -= 24;
                }

                result.endTime = autoEndTime;
                const workHours = 9; // 9시간 근무
                const breakHours = this.calculateBreakTime(workHours);
                const baseActualWorkHours = Math.max(0, workHours - breakHours);
                const annualLeaveHours = result.annualLeaveHours || 0;

                result.workHours = workHours;
                result.breakHours = breakHours;
                result.actualWorkHours = baseActualWorkHours + annualLeaveHours; // 기본근무시간 + 연차시간

                totalWorkHours += workHours;
                totalBreakHours += breakHours;

                if (this.isWeekend(record.date) || this.isHoliday(record.date)) {
                    totalOvertimeHours += result.actualWorkHours;
                } else {
                    const regularHours = Math.min(result.actualWorkHours, 8);
                    const overtimeHours = Math.max(0, result.actualWorkHours - 8);

                    totalRegularHours += regularHours;
                    totalOvertimeHours += overtimeHours;
                    workDays++;
                }

                result.status = '미완료'; // 출근만 한 상태
            } else {
                // 아직 근무하지 않은 미래 날짜인지 확인
                const today = new Date();
                const recordDate = new Date(record.date);

                if (recordDate > today && !this.isWeekend(recordDate) && !this.isHoliday(recordDate)) {
                    // 미래 평일은 근무 예정으로 처리하고 9시간(근무8시간+휴게1시간)으로 계산
                    result.status = '근무예정';
                    result.workHours = 9; // 9시간 (근무 8시간 + 휴게시간 1시간)
                    result.breakHours = 1; // 1시간 휴게시간
                    const baseActualWorkHours = 8; // 기본 근무시간 8시간
                    const annualLeaveHours = result.annualLeaveHours || 0;
                    result.actualWorkHours = baseActualWorkHours + annualLeaveHours; // 기본근무시간 + 연차시간

                    totalWorkHours += 9;
                    totalBreakHours += 1;
                    totalRegularHours += result.actualWorkHours;
                    workDays++;
                } else {
                    // 연차만 있는 경우도 처리
                    const annualLeaveHours = result.annualLeaveHours || 0;
                    result.actualWorkHours = annualLeaveHours;
                    result.status = annualLeaveHours > 0 ? '연차' : '결근';
                }
            }

            // 연차가 감지된 경우 실근무시간 재계산 (기본근무시간 + 연차시간)
            if (result.annualLeaveHours > 0) {
                let baseActualWorkHours = 0;

                // 기본 근무시간 계산 (출근/퇴근 시간이 있는 경우)
                if (result.startTime && result.endTime) {
                    const workHours = this.calculateWorkHours(result.startTime, result.endTime);
                    const breakHours = this.calculateBreakTime(workHours);
                    baseActualWorkHours = Math.max(0, workHours - breakHours);

                    // 기본 근무정보도 업데이트
                    result.workHours = workHours;
                    result.breakHours = breakHours;
                }

                // 실근무시간 = 기본근무시간 + 연차시간
                result.actualWorkHours = baseActualWorkHours + result.annualLeaveHours;

                this.addDebugLog(`${this.formatDate(record.date)} 연차 포함 실근무시간 재계산`, {
                    baseActualWorkHours: baseActualWorkHours,
                    annualLeaveHours: result.annualLeaveHours,
                    finalActualWorkHours: result.actualWorkHours
                });
            }

            // 모든 실근무시간을 합계에 누적 (0인 경우도 로그 출력)
            if (result.actualWorkHours > 0) {
                totalActualWorkHours += result.actualWorkHours;

                this.addDebugLog(`${this.formatDate(record.date)} 실근무시간 누적`, {
                    actualWorkHours: result.actualWorkHours,
                    totalActualWorkHours: totalActualWorkHours,
                    status: result.status,
                    annualLeaveHours: result.annualLeaveHours
                });
            } else {
                this.addDebugLog(`${this.formatDate(record.date)} 실근무시간 0 - 누적되지 않음`, {
                    actualWorkHours: result.actualWorkHours,
                    status: result.status,
                    startTime: result.startTime,
                    endTime: result.endTime,
                    annualLeaveHours: result.annualLeaveHours
                });
            }

            return result;
        });

        // 정상 근무일 계산 (실제로 근무해야 하는 일수)
        let normalWorkDays = 0;
        let actualWorkDays = 0; // 실제 근무한 일수

        processedData.forEach(record => {
            if (!this.isWeekend(record.date) && !this.isHoliday(record.date)) {
                normalWorkDays++; // 정상 근무일 (평일 중 공휴일이 아닌 날)

                if (record.status === '연차') {
                    // 연차는 근무한 것으로 간주하지 않음 (N7 수식에서 L7=0이면 빈값)
                } else if (record.startTime || record.endTime) {
                    actualWorkDays++; // 실제 근무
                }
            }
        });

        // N7 수식 반영: =IF(L7=0,,IF(F7="정산",8,$L7))
        // L7이 0이면 빈값, F7이 "정산"이면 8, 아니면 L7 값
        let totalRequiredHours = 0;
        processedData.forEach(record => {
            if (!this.isWeekend(record.date) && !this.isHoliday(record.date)) {
                const workHours = record.actualWorkHours || 0;

                if (workHours === 0) {
                    // N7 수식: L7=0이면 빈값 (0시간으로 처리)
                    totalRequiredHours += 0;
                } else if (record.status === '정산') {
                    // F7이 "정산"이면 8시간
                    totalRequiredHours += 8;
                } else {
                    // 그 외에는 실제 근무시간 ($L7)
                    totalRequiredHours += workHours;
                }
            }
        });

        // 필요 근무시간 = 정상 근무일 * 8
        const requiredWorkHours = normalWorkDays * 8;

        // L39: 정상 근무 시간 - 테이블에 표시된 실근무시간의 단순 합계
        let tableActualWorkHours = 0;
        processedData.forEach(record => {
            if (record.actualWorkHours > 0) {
                tableActualWorkHours += record.actualWorkHours;
            }
        });

        const normalWorkHours = tableActualWorkHours;

        // L41: 부족분 가용시간 = 정상근무시간 - 필요근무시간 (음수 허용)
        const shortageHours = normalWorkHours - requiredWorkHours;

        this.addDebugLog('부족분 계산 상세', {
            'tableActualWorkHours (테이블 실근무시간 합계)': tableActualWorkHours,
            'totalActualWorkHours (기존 실근무시간 합계)': totalActualWorkHours,
            'totalRegularHours (기존계산법)': totalRegularHours,
            'normalWorkHours (최종 사용값)': normalWorkHours,
            'requiredWorkHours (필요근무시간)': requiredWorkHours,
            'shortageHours (부족분)': shortageHours,
            '계산식': `${normalWorkHours} - ${requiredWorkHours} = ${shortageHours}`
        });

        this.calculatedResults = {
            totalWorkHours: Math.round(totalWorkHours * 100) / 100,
            totalBreakHours: Math.round(totalBreakHours * 100) / 100,
            regularHours: Math.round(tableActualWorkHours * 100) / 100,
            overtimeHours: Math.round(totalOvertimeHours * 100) / 100,
            annualLeaveDays: annualLeaveDays,
            shortageHours: Math.round(shortageHours * 100) / 100,
            normalWorkDays: normalWorkDays,
            requiredWorkHours: requiredWorkHours,
            detailedData: processedData
        };

        // 공휴일 체크 상세 분석을 위한 디버그 로그
        let weekdayCount = 0;
        let holidayCount = 0;
        let workRecordCount = 0;

        processedData.forEach(record => {
            if (!this.isWeekend(record.date)) {
                weekdayCount++;
                if (this.isHoliday(record.date)) {
                    holidayCount++;
                }
                if (record.startTime || record.endTime) {
                    workRecordCount++;
                }
            }
        });

        // 계산 과정 디버그 로그
        this.addDebugLog('부족분 계산 과정 (엑셀 공식 반영)', {
            '전체 평일수': weekdayCount,
            '공휴일수': holidayCount,
            '근무 기록 있는 일수': workRecordCount,
            '정상근무일수 (평일-공휴일)': normalWorkDays,
            '실제근무일수': actualWorkDays,
            'L39 정상근무시간': normalWorkHours,
            'L40 필요근무시간 (정상근무일*8)': requiredWorkHours,
            '연차사용일수': annualLeaveDays,
            'L41 부족분가용시간 (L39-L40)': shortageHours,
            'N7수식 총합계': totalRequiredHours
        });

        this.addDebugLog('최종 계산 결과', this.calculatedResults);
    }

    getStandardWorkDays() {
        if (this.attendanceData.length === 0) return 22;

        // 실제 데이터에서 평일(주말 제외) 일수 계산
        let totalWeekdays = 0;
        let actualHolidays = 0;

        this.attendanceData.forEach(record => {
            if (!this.isWeekend(record.date)) {
                totalWeekdays++; // 평일 총 일수

                if (this.isHoliday(record.date)) {
                    actualHolidays++; // 실제 공휴일 일수
                }
            }
        });

        const standardWorkDays = totalWeekdays - actualHolidays;

        this.addDebugLog('표준 근무일수 계산', {
            '전체 데이터 일수': this.attendanceData.length,
            '평일 총 일수': totalWeekdays,
            '공휴일 일수': actualHolidays,
            '표준 근무일수': standardWorkDays
        });

        return standardWorkDays;
    }

    displayResults() {
        const resultsSection = document.getElementById('resultsSection');
        resultsSection.style.display = 'block';

        const results = this.calculatedResults;

        // 정상 근무일 업데이트
        const normalWorkDaysElement = document.getElementById('normalWorkDays');
        if (normalWorkDaysElement) {
            normalWorkDaysElement.textContent = results.normalWorkDays;
        }

        // 부족분 가용시간 업데이트 및 색상 변경
        const shortageElement = document.getElementById('shortageHours');
        const shortageCard = document.getElementById('shortageCard');

        if (shortageElement && shortageCard) {
            const shortageHours = results.shortageHours;
            const koreanFormat = this.formatHoursToKorean(shortageHours);

            // "2시간 30분" 또는 "-1시간 30분" 형식으로 표시
            shortageElement.textContent = koreanFormat;

            // 색상 변경
            shortageCard.classList.remove('shortage', 'surplus', 'balanced');

            if (shortageHours > 0) {
                // 양수: 여유시간 있음 (초록색)
                shortageCard.classList.add('surplus');
            } else if (shortageHours < 0) {
                // 음수: 부족분 있음 (빨간색)
                shortageCard.classList.add('shortage');
            } else {
                // 0: 정확히 맞음 (회색)
                shortageCard.classList.add('balanced');
            }

            this.addDebugLog('부족분 UI 업데이트', {
                shortageHours: shortageHours,
                koreanFormat: koreanFormat,
                colorClass: shortageHours > 0 ? 'surplus' : shortageHours < 0 ? 'shortage' : 'balanced'
            });
        }

        // 디버깅용 정상 근무시간 및 필요 근무시간 업데이트
        const normalWorkHoursElement = document.getElementById('normalWorkHours');
        const requiredWorkHoursElement = document.getElementById('requiredWorkHours');

        if (normalWorkHoursElement) {
            normalWorkHoursElement.textContent = results.regularHours;
        }

        if (requiredWorkHoursElement) {
            requiredWorkHoursElement.textContent = results.requiredWorkHours;
        }

        this.displayDetailedTable(results.detailedData);

        if (this.attendanceData.length > 0) {
            resultsSection.scrollIntoView({ behavior: 'smooth' });
        }
    }

    updateResultsOnly() {
        // 테이블을 다시 그리지 않고 숫자만 업데이트
        const results = this.calculatedResults;

        // 정상 근무일 업데이트
        const normalWorkDaysElement = document.getElementById('normalWorkDays');
        if (normalWorkDaysElement) {
            normalWorkDaysElement.textContent = results.normalWorkDays;
        }

        // 부족분 가용시간 업데이트 및 색상 변경
        const shortageElement = document.getElementById('shortageHours');
        const shortageCard = document.getElementById('shortageCard');

        if (shortageElement && shortageCard) {
            const shortageHours = results.shortageHours;
            const koreanFormat = this.formatHoursToKorean(shortageHours);

            // "2시간 30분" 또는 "-1시간 30분" 형식으로 표시
            shortageElement.textContent = koreanFormat;

            // 색상 변경
            shortageCard.classList.remove('shortage', 'surplus', 'balanced');

            if (shortageHours > 0) {
                shortageCard.classList.add('surplus');
            } else if (shortageHours < 0) {
                shortageCard.classList.add('shortage');
            } else {
                shortageCard.classList.add('balanced');
            }

            this.addDebugLog('부족분만 업데이트', {
                shortageHours: shortageHours,
                koreanFormat: koreanFormat,
                normalWorkDays: results.normalWorkDays,
                colorClass: shortageHours > 0 ? 'surplus' : shortageHours < 0 ? 'shortage' : 'balanced'
            });
        }

        // 디버깅용 정상 근무시간 및 필요 근무시간 업데이트
        const normalWorkHoursElement = document.getElementById('normalWorkHours');
        const requiredWorkHoursElement = document.getElementById('requiredWorkHours');

        if (normalWorkHoursElement) {
            normalWorkHoursElement.textContent = results.regularHours;
        }

        if (requiredWorkHoursElement) {
            requiredWorkHoursElement.textContent = results.requiredWorkHours;
        }
    }

    displayDetailedTable(data) {
        const tableBody = document.getElementById('detailTableBody');
        tableBody.innerHTML = '';

        data.forEach((record, index) => {
            const row = tableBody.insertRow();

            // 날짜
            row.insertCell(0).textContent = this.formatDate(record.date);

            // 편집 가능한 출근시간
            const startTimeCell = row.insertCell(1);
            const startTimeInput = document.createElement('input');
            startTimeInput.type = 'text';
            startTimeInput.value = this.formatTime(record.startTime);
            startTimeInput.className = 'time-input';
            startTimeInput.placeholder = 'HH:MM';
            startTimeInput.addEventListener('blur', () => this.updateTimeAndRecalculate(index, 'startTime', startTimeInput.value));
            startTimeInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    startTimeInput.blur();
                }
            });
            startTimeCell.appendChild(startTimeInput);

            // 편집 가능한 퇴근시간
            const endTimeCell = row.insertCell(2);
            const endTimeInput = document.createElement('input');
            endTimeInput.type = 'text';
            endTimeInput.value = this.formatTime(record.endTime);
            endTimeInput.className = 'time-input';
            endTimeInput.placeholder = 'HH:MM';
            endTimeInput.addEventListener('blur', () => this.updateTimeAndRecalculate(index, 'endTime', endTimeInput.value));
            endTimeInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    endTimeInput.blur();
                }
            });
            endTimeCell.appendChild(endTimeInput);

            row.insertCell(3).textContent = record.workHours > 0 ? `${record.workHours.toFixed(2)}시간` : '-';
            row.insertCell(4).textContent = record.breakHours > 0 ? `${record.breakHours.toFixed(2)}시간` : '-';

            // 편집 가능한 연차시간 (5번째 컬럼)
            const annualLeaveCell = row.insertCell(5);
            const annualLeaveInput = document.createElement('input');
            annualLeaveInput.type = 'text';
            // 데이터에서 연차가 감지된 경우 자동 기입
            let annualLeaveHours = record.annualLeaveHours || 0;
            annualLeaveInput.value = annualLeaveHours > 0 ? annualLeaveHours : '0';
            annualLeaveInput.className = 'annual-leave-input';
            annualLeaveInput.placeholder = '0';
            annualLeaveInput.addEventListener('blur', () => this.updateAnnualLeaveAndRecalculate(index, annualLeaveInput.value));
            annualLeaveInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    annualLeaveInput.blur();
                }
            });
            annualLeaveCell.appendChild(annualLeaveInput);

            // 실근무시간 (6번째 컬럼)
            row.insertCell(6).textContent = record.actualWorkHours > 0 ? `${record.actualWorkHours.toFixed(2)}시간` : '-';

            // 공휴일 체크박스 셀 추가
            const holidayCell = row.insertCell(7);
            if (!this.isWeekend(record.date)) { // 평일에만 체크박스 표시
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.checked = this.isHoliday(record.date);
                checkbox.addEventListener('change', (e) => {
                    const dateStr = record.date.toDateString();
                    if (e.target.checked) {
                        this.holidays.add(dateStr);
                    } else {
                        // 기본 공휴일이 아닌 경우에만 제거
                        if (!this.isDefaultHoliday(record.date)) {
                            this.holidays.delete(dateStr);
                        } else {
                            // 기본 공휴일인 경우 체크를 다시 켬
                            e.target.checked = true;
                            alert('기본 공휴일은 해제할 수 없습니다.');
                            return;
                        }
                    }
                    // 재계산 및 부족분만 업데이트 (테이블 다시 그리지 않음)
                    this.calculateAttendance();
                    this.updateResultsOnly();
                    this.addDebugLog('공휴일 체크박스 변경으로 인한 재계산', {
                        dateChanged: this.formatDate(record.date),
                        isNowHoliday: e.target.checked,
                        totalHolidays: this.holidays.size
                    });
                });
                holidayCell.appendChild(checkbox);
            }

            const statusCell = row.insertCell(8);

            // 비고란에 '근무예정'만 표시하고 나머지는 비워둠
            if (record.status === '근무예정') {
                statusCell.textContent = '근무예정';
            } else {
                statusCell.textContent = '';
            }

            // 토요일, 일요일 배경 회색 처리
            if (this.isWeekend(record.date)) {
                row.style.backgroundColor = '#f5f5f5';
                row.style.color = '#666';
            }

            // 출근시간만 있고 퇴근시간이 자동계산된 경우 퇴근시간만 빨간색 처리
            if (record.status === '미완료') {
                const endTimeCell = row.cells[2]; // 퇴근시간 셀
                endTimeCell.style.color = '#c62828';
                endTimeCell.style.fontWeight = 'bold';
            }

            // 상태별 클래스 적용
            if (record.status === '미완료') {
                statusCell.className = 'status-late';
            } else if (record.status.includes('연차') || record.status.includes('휴가')) {
                statusCell.className = 'status-leave';
            } else if (record.status === '근무예정') {
                statusCell.className = 'status-normal';
            } else {
                statusCell.className = 'status-normal';
            }
        });
    }

    updateTimeAndRecalculate(index, field, timeString) {
        // 시간 파싱
        const parsedTime = this.parseTimeString(timeString);

        if (parsedTime) {
            // 데이터 업데이트
            this.calculatedResults.detailedData[index][field] = parsedTime;

            // 해당 레코드 재계산
            this.recalculateRecord(index);

            // 전체 합계 재계산 및 UI 업데이트
            this.recalculateAndUpdateResults();
        }
    }

    updateAnnualLeaveAndRecalculate(index, annualLeaveString) {
        const record = this.calculatedResults.detailedData[index];
        const annualLeaveHours = parseFloat(annualLeaveString.trim()) || 0;

        // 연차시간을 별도 프로퍼티로 저장
        record.annualLeaveHours = annualLeaveHours;

        // 기본 근무시간 재계산
        this.recalculateRecord(index);

        // 전체 합계 재계산 및 UI 업데이트
        this.recalculateAndUpdateResults();
    }

    parseTimeString(timeString) {
        if (!timeString || timeString.trim() === '' || timeString === '-') {
            return null;
        }

        const timeStr = timeString.trim();
        const timePattern = /^(\d{1,2}):(\d{2})$/;
        const match = timeStr.match(timePattern);

        if (match) {
            const hours = parseInt(match[1]);
            const minutes = parseInt(match[2]);
            if (hours >= 0 && hours < 24 && minutes >= 0 && minutes < 60) {
                return { hours, minutes };
            }
        }

        return null;
    }

    recalculateRecord(index) {
        const record = this.calculatedResults.detailedData[index];
        const annualLeaveHours = record.annualLeaveHours || 0;

        // 기본 근무시간 계산
        let baseActualWorkHours = 0;
        if (record.startTime && record.endTime) {
            const workHours = this.calculateWorkHours(record.startTime, record.endTime);
            const breakHours = this.calculateBreakTime(workHours);
            baseActualWorkHours = Math.max(0, workHours - breakHours);

            record.workHours = workHours;
            record.breakHours = breakHours;
            record.status = '정상';
        } else {
            record.workHours = 0;
            record.breakHours = 0;
            record.status = '결근';
        }

        // 실근무시간 = 기본근무시간(휴게시간 제외) + 연차시간
        record.actualWorkHours = baseActualWorkHours + annualLeaveHours;

        this.addDebugLog(`${this.formatDate(record.date)} 재계산`, {
            baseActualWorkHours: baseActualWorkHours,
            annualLeaveHours: annualLeaveHours,
            finalActualWorkHours: record.actualWorkHours
        });
    }

    recalculateAndUpdateResults() {
        // 테이블에 표시된 실근무시간의 합계 계산
        let tableActualWorkHours = 0;
        this.calculatedResults.detailedData.forEach(record => {
            if (record.actualWorkHours > 0) {
                tableActualWorkHours += record.actualWorkHours;
            }
        });

        // 결과 업데이트
        this.calculatedResults.regularHours = Math.round(tableActualWorkHours * 100) / 100;

        // 부족분 재계산
        const normalWorkDays = this.calculatedResults.normalWorkDays;
        const requiredWorkHours = normalWorkDays * 8;
        const shortageHours = tableActualWorkHours - requiredWorkHours;
        this.calculatedResults.shortageHours = Math.round(shortageHours * 100) / 100;

        // UI만 업데이트 (테이블은 다시 그리지 않음)
        this.updateResultsOnly();
        this.updateTableValues();
    }

    updateTableValues() {
        // 테이블의 계산된 값들을 업데이트
        const rows = document.getElementById('detailTableBody').rows;

        this.calculatedResults.detailedData.forEach((record, index) => {
            if (rows[index]) {
                const row = rows[index];
                // 근무시간 (3번째 컬럼)
                row.cells[3].textContent = record.workHours > 0 ? `${record.workHours.toFixed(2)}시간` : '-';
                // 휴게시간 (4번째 컬럼)
                row.cells[4].textContent = record.breakHours > 0 ? `${record.breakHours.toFixed(2)}시간` : '-';
                // 연차시간은 input이므로 업데이트하지 않음 (5번째 컬럼)
                // 실근무시간 (6번째 컬럼)
                row.cells[6].textContent = record.actualWorkHours > 0 ? `${record.actualWorkHours.toFixed(2)}시간` : '-';
                // 비고 (8번째 컬럼)
                if (record.status === '근무예정') {
                    row.cells[8].textContent = '근무예정';
                } else {
                    row.cells[8].textContent = '';
                }
            }
        });
    }

    formatDate(date) {
        if (!date) return '-';
        const weekDays = ['일', '월', '화', '수', '목', '금', '토'];
        const dayOfWeek = weekDays[date.getDay()];

        const dateStr = date.toLocaleDateString('ko-KR', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });

        return `${dateStr} (${dayOfWeek})`;
    }

    formatTime(time) {
        if (!time) return '-';
        return `${time.hours.toString().padStart(2, '0')}:${time.minutes.toString().padStart(2, '0')}`;
    }

    formatHoursToHHMM(hours) {
        if (!hours && hours !== 0) return '00:00';

        const absHours = Math.abs(hours);
        const wholeHours = Math.floor(absHours);
        const minutes = Math.round((absHours - wholeHours) * 60);

        const hhMM = `${wholeHours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;

        // 음수인 경우 - 부호 추가
        return hours < 0 ? `-${hhMM}` : hhMM;
    }

    formatHoursToKorean(hours) {
        if (!hours && hours !== 0) return '0시간 0분';

        const absHours = Math.abs(hours);
        const wholeHours = Math.floor(absHours);
        const minutes = Math.round((absHours - wholeHours) * 60);

        let result = '';
        if (wholeHours > 0) {
            result += `${wholeHours}시간`;
        }
        if (minutes > 0) {
            if (result) result += ' ';
            result += `${minutes}분`;
        }

        // 둘 다 0인 경우
        if (!result) {
            result = '0시간 0분';
        }

        // 음수인 경우 - 부호 추가
        return hours < 0 ? `-${result}` : result;
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new AttendanceCalculator();
});