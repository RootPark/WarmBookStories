import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, VerticalAlign, AlignmentType, HeadingLevel } from 'docx';
import mammoth from 'mammoth';

const FileButtons = () => {
  const [file1, setFile1] = useState(null);
  const [jsonData, setJsonData] = useState(null);
  const [selectedDate, setSelectedDate] = useState('');
  const [bestAssignee, setBestAssignee] = useState('');
  const [participantsFile, setParticipantsFile] = useState(null);
  const [assignmentFile, setAssignmentFile] = useState(null);
  const [assignmentFinalFile, setAssignmentFinalFile] = useState(null);

  const handleFile1Change = (event) => {
    const file = event.target.files[0];
    if (file && file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      setFile1(file);
    } else {
      alert('Please select an XLSX file');
    }
  };

  const handleDateChange = (event) => {
    setSelectedDate(event.target.value);
  };

  const convertExcelToJson = async (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const data = e.target.result;
          const workbook = XLSX.read(data, { type: 'binary' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: 1 });
          
          const members = {};
          jsonData.forEach(row => {
            if (row['성명']) {
              members[row['성명']] = row;
            }
          });

          const jsonString = JSON.stringify(members, null, 4);
          const blob = new Blob([jsonString], { type: 'application/json' });
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'members.json';
          a.click();
          window.URL.revokeObjectURL(url);

          setJsonData(members);
          resolve(members);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = (error) => reject(error);
      reader.readAsBinaryString(file);
    });
  };

  const handleButton1Click = async () => {
    if (!file1) {
      alert('Please select an Excel file first');
      return;
    }

    try {
      await convertExcelToJson(file1);
      console.log('Conversion completed and file downloaded');
    } catch (error) {
      console.error('Error converting file:', error);
      alert('Error converting file. Please check the console for details.');
    }
  };

  const extractNames = (line) => {
    const idx = line.lastIndexOf(':');
    const contents = line.substring(idx + 1).trim();
    const names = contents.split(/[\s,/]+/)
      .filter(x => x && !/[0-9][0-9]기/.test(x) && !/^[a-zA-Z]$/.test(x));
    return names.sort();
  };

  const getAssignees = (pinfo, membersData) => {
    const assignees = [];
    const graduates = [];

    pinfo.participants.forEach(name => {
      if (membersData[name]) {
        if (membersData[name]["회원 등급"] === "일반" && 
            !pinfo.presenters.includes(name)) {
          assignees.push(name);
        } else if (membersData[name]["회원 등급"] === "수료") {
          graduates.push(name);
        }
      }
    });

    return {
      ...pinfo,
      assignees: assignees.sort(),
      graduates: graduates.sort()
    };
  };

  const processDocxFiles = async (files) => {
    const fileContents = await Promise.all(files.map(async file => {
      const arrayBuffer = await file.file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer });
      return { type: file.type, content: result.value };
    }));

    // participants 파일 처리
    const participantsFile = fileContents.find(f => f.type === 'participants');
    if (!participantsFile) {
      throw new Error('서기록 파일을 찾을 수 없습니다.');
    }

    const info = {
      participants: [],
      latecomers_a: [],
      latecomers_b: [],
      latecomers_c: [],
      non_voters: [],
      absentees_a: [],
      absentees_b: [],
      presenters: []
    };

    // 서기록 파일에서 정보 추출
    const lines = participantsFile.content.split('\n');
    lines.forEach(line => {
      if (line.includes('3:35')) {
        info.latecomers_a = extractNames(line);
      } else if (line.includes('4:00')) {
        info.latecomers_b = extractNames(line);
      } else if (line.includes('4:30')) {
        info.latecomers_c = extractNames(line);
      } else if (line.includes('미투표자')) {
        info.non_voters = extractNames(line);
      } else if (line.includes('무단')) {
        info.absentees_b = extractNames(line);
      } else if (line.includes('정오 이후')) {
        info.absentees_a = extractNames(line);
      } else if (line.includes('조:')) {
        info.participants.push(...extractNames(line));
      } else if (line.includes('발제자:')) {
        info.presenters = extractNames(line);
      }
    });

    // 과제 파일 처리
    const assignmentFile = fileContents.find(f => f.type === 'assignment');
    const assignmentFinalFile = fileContents.find(f => f.type === 'assignment_final');
    if (!assignmentFile || !assignmentFinalFile) {
      throw new Error('과제 파일을 찾을 수 없습니다.');
    }

    // 과제 테이블 찾기
    const findTables = (content) => {
      const tables = [];
      const lines = content.split('\n');
      let currentTable = [];
      
      lines.forEach(line => {
        if (line.includes('\t')) {
          currentTable.push(line.split('\t'));
        } else if (currentTable.length > 0) {
          if (currentTable[0].length === 2) {
            tables.push(currentTable);
          }
          currentTable = [];
        }
      });
      
      return tables;
    };

    const assignmentTables = findTables(assignmentFile.content);
    const assignmentFinalTables = findTables(assignmentFinalFile.content);

    // 과제 제출 상태 확인
    const countNames = (tables, assignees) => {
      const namesCount = {};
      tables.forEach(table => {
        table.forEach(row => {
          const name = row[0].trim();
          const answer = row[1].trim();
          if (name && answer && assignees.includes(name)) {
            namesCount[name] = (namesCount[name] || 0) + 1;
          }
        });
      });
      return namesCount;
    };

    const processedInfo = getAssignees(info, jsonData);
    const namesCount = countNames(assignmentTables, processedInfo.assignees);
    const namesCountFinal = countNames(assignmentFinalTables, processedInfo.assignees);

    // 과제 제출 상태 분류
    const failers = [];
    const lateSubmitters = [];
    const completers = [];

    processedInfo.assignees.forEach(name => {
      if (namesCountFinal[name] < 3) {
        failers.push(name);
      } else if (namesCount[name] < 3 && namesCountFinal[name] >= 3) {
        lateSubmitters.push(name);
      } else {
        completers.push(name);
      }
    });

    processedInfo.failers = failers.sort();
    processedInfo.late_submitters = lateSubmitters.sort();
    processedInfo.completers = completers.sort();
    processedInfo.best_assignee = bestAssignee;

    // 회계록 문서 생성
    return createAccountDoc(processedInfo);
  };

  const downloadDocument = async (doc, filename) => {
    const blob = await Packer.toBlob(doc);
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const createAccountDoc = (info) => {
    const calculateFine = (list, amount) => (list || []).length * amount;
    
    const fines = {
      non_voters: calculateFine(info.non_voters, 1000),
      absentees_a: calculateFine(info.absentees_a, 7000),
      absentees_b: calculateFine(info.absentees_b, 20000),
      latecomers_a: calculateFine(info.latecomers_a, 2000),
      latecomers_b: calculateFine(info.latecomers_b, 4000),
      latecomers_c: calculateFine(info.latecomers_c, 6000),
      late_submitters: calculateFine(info.late_submitters, 2000),
      failers: calculateFine(info.failers, 4000)
    };

    const totalFine = Object.values(fines).reduce((a, b) => a + b, 0);

    // 표 스타일 설정
    const tableStyle = {
      width: {
        size: 9000,
        type: WidthType.DXA,
      },
      margins: {
        top: 100,
        bottom: 100,
        left: 100,
        right: 100,
      },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1 },
        bottom: { style: BorderStyle.SINGLE, size: 1 },
        left: { style: BorderStyle.SINGLE, size: 1 },
        right: { style: BorderStyle.SINGLE, size: 1 },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
        insideVertical: { style: BorderStyle.SINGLE, size: 1 },
      },
    };

    // 셀 스타일 설정
    const cellStyle = {
      margins: {
        top: 100,
        bottom: 100,
        left: 100,
        right: 100,
      },
      verticalAlign: VerticalAlign.CENTER,
    };

    // 첫 번째 표 생성 (벌금 정보)
    const mainTable = new Table({
      ...tableStyle,
      rows: [
        // 헤더 행
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              width: { size: 2000, type: WidthType.DXA },
              children: [new Paragraph({ text: "구분", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              width: { size: 4000, type: WidthType.DXA },
              children: [new Paragraph({ text: "이름", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              width: { size: 1500, type: WidthType.DXA },
              children: [new Paragraph({ text: "벌금 구분", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              width: { size: 1500, type: WidthType.DXA },
              children: [new Paragraph({ text: "벌금", alignment: AlignmentType.CENTER })],
            }),
          ],
        }),
        // 미투표자
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "미투표자", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: info.non_voters?.join(', ') || '' })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "1,000원", alignment: AlignmentType.RIGHT })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: `${fines.non_voters}원`, alignment: AlignmentType.RIGHT })],
            }),
          ],
        }),
        // 불참자(정오 이후)
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "불참자(정오 이후)", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: info.absentees_a?.join(', ') || '' })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "7,000원", alignment: AlignmentType.RIGHT })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: `${fines.absentees_a}원`, alignment: AlignmentType.RIGHT })],
            }),
          ],
        }),
        // 무단 불참자
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "무단 불참자", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: info.absentees_b?.join(', ') || '' })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "20,000원", alignment: AlignmentType.RIGHT })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: `${fines.absentees_b}원`, alignment: AlignmentType.RIGHT })],
            }),
          ],
        }),
        // 지각자
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "지각자", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              children: [
                new Paragraph({ text: "3:35~3:59 " + (info.latecomers_a?.join(', ') || '') }),
                new Paragraph({ text: "4:00~4:29 " + (info.latecomers_b?.join(', ') || '') }),
                new Paragraph({ text: "4:30~ " + (info.latecomers_c?.join(', ') || '') }),
              ],
            }),
            new TableCell({
              ...cellStyle,
              children: [
                new Paragraph({ text: "2,000원", alignment: AlignmentType.RIGHT }),
                new Paragraph({ text: "4,000원", alignment: AlignmentType.RIGHT }),
                new Paragraph({ text: "6,000원", alignment: AlignmentType.RIGHT }),
              ],
            }),
            new TableCell({
              ...cellStyle,
              children: [
                new Paragraph({ text: `${fines.latecomers_a}원`, alignment: AlignmentType.RIGHT }),
                new Paragraph({ text: `${fines.latecomers_b}원`, alignment: AlignmentType.RIGHT }),
                new Paragraph({ text: `${fines.latecomers_c}원`, alignment: AlignmentType.RIGHT }),
              ],
            }),
          ],
        }),
        // 과제 지각자
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "과제 지각자", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: info.late_submitters?.join(', ') || '' })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "2,000원", alignment: AlignmentType.RIGHT })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: `${fines.late_submitters}원`, alignment: AlignmentType.RIGHT })],
            }),
          ],
        }),
        // 과제 미제출자
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "과제 미제출자", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: info.failers?.join(', ') || '' })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "4,000원", alignment: AlignmentType.RIGHT })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: `${fines.failers}원`, alignment: AlignmentType.RIGHT })],
            }),
          ],
        }),
        // 총 벌금
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "총 벌금", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              columnSpan: 3,
              children: [new Paragraph({ text: `${totalFine}원`, alignment: AlignmentType.RIGHT })],
            }),
          ],
        }),
      ],
    });

    // 두 번째 표 생성 (포상금 정보)
    const rewardTable = new Table({
      ...tableStyle,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              width: { size: 2000, type: WidthType.DXA },
              children: [new Paragraph({ text: "발제자 포상금(500원)", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              width: { size: 7000, type: WidthType.DXA },
              children: [new Paragraph({ text: info.presenters?.join(', ') || '' })],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: "우수과제자", alignment: AlignmentType.CENTER })],
            }),
            new TableCell({
              ...cellStyle,
              children: [new Paragraph({ text: info.best_assignee || '' })],
            }),
          ],
        }),
      ],
    });

    return new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: `따책회계록 - ${selectedDate}`,
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 200 },
          }),
          mainTable,
          new Paragraph({ spacing: { before: 200, after: 200 } }),
          rewardTable,
        ],
      }],
    });
  };

  const handleDocxFileChange = (event, setFile) => {
    const file = event.target.files[0];
    if (file && file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      setFile(file);
    } else {
      alert('Please select a DOCX file');
    }
  };

  const handleButton2Click = async () => {
    if (!selectedDate || !jsonData || !bestAssignee || !participantsFile || !assignmentFile || !assignmentFinalFile) {
      alert('모든 파일과 정보를 입력해주세요.');
      return;
    }

    try {
      const files = [
        { file: participantsFile, type: 'participants' },
        { file: assignmentFile, type: 'assignment' },
        { file: assignmentFinalFile, type: 'assignment_final' }
      ];

      const doc = await processDocxFiles(files);
      await downloadDocument(doc, `따책회계록_${selectedDate}.docx`);

    } catch (error) {
      console.error('Error processing files:', error);
      alert('문서 처리 중 오류가 발생했습니다.');
    }
  };

  return (
    <div className="file-buttons-container">
      <div className="file-inputs">
        <div>
          <input
            type="file"
            accept=".xlsx"
            onChange={handleFile1Change}
            id="file1"
            style={{ display: 'none' }}
          />
          <label htmlFor="file1" className="file-button">
            Select Excel File
          </label>
          <span>{file1 ? file1.name : 'No file selected'}</span>
        </div>
        
        <div>
          <input
            type="date"
            value={selectedDate}
            onChange={handleDateChange}
            className="date-input"
          />
        </div>

        <div>
          <input
            type="text"
            value={bestAssignee}
            onChange={(e) => setBestAssignee(e.target.value)}
            placeholder="우수과제자 이름"
            className="text-input"
          />
        </div>

        <div>
          <input
            type="file"
            accept=".docx"
            onChange={(e) => handleDocxFileChange(e, setParticipantsFile)}
            id="participantsFile"
            style={{ display: 'none' }}
          />
          <label htmlFor="participantsFile" className="file-button">
            서기록 파일 넣기
          </label>
          <span>{participantsFile ? participantsFile.name : 'No file selected'}</span>
        </div>

        <div>
          <input
            type="file"
            accept=".docx"
            onChange={(e) => handleDocxFileChange(e, setAssignmentFile)}
            id="assignmentFile"
            style={{ display: 'none' }}
          />
          <label htmlFor="assignmentFile" className="file-button">
            정오에 다운받은 과제 파일 넣기
          </label>
          <span>{assignmentFile ? assignmentFile.name : 'No file selected'}</span>
        </div>

        <div>
          <input
            type="file"
            accept=".docx"
            onChange={(e) => handleDocxFileChange(e, setAssignmentFinalFile)}
            id="assignmentFinalFile"
            style={{ display: 'none' }}
          />
          <label htmlFor="assignmentFinalFile" className="file-button">
            최종 과제 파일 넣기
          </label>
          <span>{assignmentFinalFile ? assignmentFinalFile.name : 'No file selected'}</span>
        </div>
      </div>

      <div className="action-buttons">
        <button onClick={handleButton1Click}>JSON 생성</button>
        <button onClick={handleButton2Click}>문서 생성</button>
      </div>

      {jsonData && (
        <div className="json-viewer">
          <h3>JSON 데이터 뷰어</h3>
          <pre>{JSON.stringify(jsonData, null, 2)}</pre>
        </div>
      )}
    </div>
  );
};

export default FileButtons; 