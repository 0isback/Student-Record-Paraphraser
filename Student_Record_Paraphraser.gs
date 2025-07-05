function createBatchPrompt(batch) {
  const numbered = batch.map((item, i) => `${i + 1}. ${item.text}`).join("\n");
  return `다음 문장들을 의미는 유지하되 사용하는 어휘나 어구를 바꿔서 자연스럽게 다시 써줘. 원래 문장들과 같이 '~함', '~해봄', '~나눔'처럼 명사형 종결어미, 흔히 말하는 '음슴체'로 바꿔줘. 번호별로 결과를 출력해줘:\n\n${numbered}`;
}

function batchParaphrase(prompt, count) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"); 
  // 개인 OPENAI_API_KEY 를 Apps Script '스크립트 속성'으로 추가해서 사용하면 됩니다.

  if (!apiKey) throw new Error("OpenAI API 키가 없습니다.");

  const payload = {
    model: "gpt-3.5-turbo",
    messages: [{ role: "user", content: prompt }],
    temperature: 0.8,
    max_tokens: 2048,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload),
  };

  try {
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    const json = JSON.parse(response.getContentText());
    const reply = json.choices?.[0]?.message?.content;

    const lines = reply.split(/\n+/).filter(line => /^\d+\.\s/.test(line));
    return lines.map(line => line.replace(/^\d+\.\s*/, "").trim()).slice(0, count);
  } catch (e) {
    Logger.log("GPT 요청 오류: " + e);
    return null;
  }
}

function generateProgressBar(current, total, width) {
  const filled = Math.round((current / total) * width);
  return `[${"█".repeat(filled)}${"▒".repeat(width - filled)}]`;
}

function onParaphraseButtonClick() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const inputRange = sheet.getRange("F3:F52");
  const outputRange = sheet.getRange("H3:H52");
  const statusCell = sheet.getRange("G1");

  // 초기화
  outputRange.clearContent();
  statusCell.setValue("처리 시작...");
  SpreadsheetApp.flush();

  const inputValues = inputRange.getValues().map(row => row[0]);
  const validRows = inputValues
    .map((text, i) => ({ index: i, text }))
    .filter(item => item.text && item.text.trim() !== "");

  const total = validRows.length;
  if (total === 0) {
    statusCell.setValue("처리할 문장이 없습니다.");
    return;
  }

  const batchSize = 10;
  let completed = 0;
  const startTime = new Date();

  for (let i = 0; i < total; i += batchSize) {
    const batch = validRows.slice(i, i + batchSize);
    const prompt = createBatchPrompt(batch);
    const paraphrasedList = batchParaphrase(prompt, batch.length);

    if (!paraphrasedList) {
      statusCell.setValue(`오류 발생: ${completed}/${total} 문장 처리됨`);
      SpreadsheetApp.flush();
      return;
    }

    // 결과 쓰기
    for (let j = 0; j < batch.length; j++) {
      const row = batch[j].index + 1;
      outputRange.getCell(row, 1).setValue(paraphrasedList[j]);
    }

    completed += batch.length;

    // 진행률 표시
    const elapsedMs = new Date() - startTime;
    const avgTime = elapsedMs / completed;
    const remaining = total - completed;
    const estSec = Math.round((remaining * avgTime) / 1000);
    const progressBar = generateProgressBar(completed, total, 20);

    statusCell.setValue(
      `진행 중: ${completed}/${total} 완료됨 ${progressBar} (남은 예상 시간: ${estSec}초)`
    );
    SpreadsheetApp.flush();

    // 자동 스크롤 + 셀 강조
    if (completed % 10 === 0) {
      const nextRow = Math.min(3 + completed, 52);
      const targetCell = sheet.getRange(`H${nextRow}`);
      targetCell.setBackground("#fff59d");
      sheet.setActiveRange(targetCell); // 자동 스크롤
      SpreadsheetApp.flush();
      Utilities.sleep(300);
      targetCell.setBackground(null);
    }
  }

  // 완료 메시지
  statusCell.setValue("전체 문장 처리 완료!");
  SpreadsheetApp.flush();
  Utilities.sleep(2000);
  statusCell.clearContent();

  // 맨 위로 커서 복귀
  sheet.setActiveRange(sheet.getRange("H3"));
}
