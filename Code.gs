const SHEET_NAME = 'Templates';

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Clinical Narrative Generator | LinDoc')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  let ss;

  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create('Clinical_Narrative_Templates_DB');
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
  }
  sheet.clear();

  const rows = [
    ['id', 'name', 'description', 'body'],
    [
      'standard',
      '1. 標準精煉型 (Standard)',
      '日常病房 Admission / 交班基準版',
      'A {{age}}-year-old {{sex}} with a history of {{comorbidities}}, {{baseline_status}}, and {{surgical_hx}} {{onset_verb}} {{chief_complaint}} approximately {{duration}} before admission. Despite {{prior_care}}, progressive {{worsening_sx}} prompted ED presentation on {{ed_date}}, followed by ward admission on {{admit_date}} for {{working_dx}}.\n\nInitial evaluation demonstrated {{lab_findings}} and {{organ_dysfunction}}. Imaging with {{imaging_modality}} revealed {{imaging_findings}}. The patient remained {{stability_status}} without clinical signs of {{pertinent_negatives}}. Following discussion with the {{consult_team}} team, {{decision_intervention}}. Follow-up {{specimen}} obtained on {{micro_date}} detected {{pathogen}}, confirming {{final_dx}} and prompting transfer to the {{dest_service}} service for specialized management.'
    ],
    [
      'academic',
      '2. 高密度專科術語型 (Academic)',
      '感染症/重症會診、學術病例報告版',
      'A {{age}}-year-old {{sex}} with underlying {{comorbidities}}, {{baseline_status}}, and {{surgical_hx}} experienced subacute onset of {{chief_complaint}} over the preceding {{duration}} prior to presentation. Refractory to {{prior_care}}, clinical deterioration characterized by {{worsening_sx}} necessitated emergency department evaluation on {{ed_date}}.\n\nInitial diagnostic workup was remarkable for {{lab_findings}} concomitant with {{organ_dysfunction}}. Cross-sectional imaging via {{imaging_modality}} delineated {{imaging_findings}}. Hemodynamically, the patient was {{stability_status}} with pertinent negatives including absence of {{pertinent_negatives}}. In consultation with {{consult_team}}, a consensus was reached to execute {{decision_intervention}}. Subsequent microbiological surveillance of {{specimen}} collected on {{micro_date}} yielded {{pathogen}}, substantiating the diagnosis of {{final_dx}} and warranting admission to the {{dest_service}} service on {{admit_date}} for definitive antimicrobial stewardship.'
    ],
    [
      'telegraphic',
      '3. 極簡電報式摘要 (Telegraphic)',
      '急診快速收治 / SBAR 床邊交班版',
      '{{age}}-year-old {{sex}} with {{comorbidities}}, {{baseline_status}}, and {{surgical_hx}} presented on {{ed_date}} with {{duration}} of {{chief_complaint}} refractory to {{prior_care}} ({{worsening_sx}}).\n\nED workup showed {{lab_findings}}, {{organ_dysfunction}}, and {{imaging_modality}} revealing {{imaging_findings}}; {{stability_status}} without {{pertinent_negatives}}. Following {{consult_team}} consultation, {{decision_intervention}}. {{specimen}} from {{micro_date}} yielded {{pathogen}}, confirming {{final_dx}}; admitted/transferred to {{dest_service}} on {{admit_date}} for {{working_dx}}.'
    ],
    [
      'narrative',
      '4. 敘事文學修辭型 (Narrative)',
      'Grand Rounds 病例開場 / 人文病歷版',
      'For a {{age}}-year-old {{sex}} whose medical background was already shadowed by {{comorbidities}}, {{baseline_status}}, and {{surgical_hx}}, the onset of {{chief_complaint}} roughly {{duration}} before admission marked the beginning of a steady clinical decline. Though {{prior_care}} was attempted, worsening {{worsening_sx}} finally drove the patient to the emergency department on {{ed_date}}.\n\nUpon arrival, laboratory studies uncovered {{lab_findings}} alongside {{organ_dysfunction}}, while {{imaging_modality}} exposed {{imaging_findings}}. Yet amidst these derangements, the patient remained {{stability_status}}, spared from {{pertinent_negatives}}. Weighing the risks, the medical and {{consult_team}} teams chose the path of {{decision_intervention}}. The decisive clue emerged on {{micro_date}}, when {{specimen}} revealed {{pathogen}}—unmasking {{final_dx}} and leading to transfer to the {{dest_service}} service on {{admit_date}} for {{working_dx}}.'
    ],
    [
      'plain',
      '5. 通俗平實跨科溝通型 (Plain)',
      '跨職類團隊 / 家屬病情摘要版',
      'This is a {{age}}-year-old {{sex}} with a history of {{comorbidities}} ({{baseline_status}}) and {{surgical_hx}}. About {{duration}} before coming to the hospital, the patient developed {{chief_complaint}}. Even though the patient received {{prior_care}}, the {{worsening_sx}} continued to get worse, leading to an emergency room visit on {{ed_date}} and hospital admission on {{admit_date}} for {{working_dx}}.\n\nEarly hospital tests showed {{lab_findings}} and {{organ_dysfunction}}. A {{imaging_modality}} showed {{imaging_findings}}. At that time, the patient was {{stability_status}} and showed no signs of {{pertinent_negatives}}. After talking with the {{consult_team}} doctors, the team proceeded with {{decision_intervention}}. Later test results from {{micro_date}} found {{pathogen}} in the {{specimen}}, confirming {{final_dx}}. The patient was then moved to the {{dest_service}} team for focused treatment.'
    ]
  ];

  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.setFrozenRows(1);
  Logger.log('Database initialized successfully! Sheet URL: ' + ss.getUrl());
}

function getTemplates() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) return [];

  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const templates = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      templates.push({
        id: data[i][0].toString().trim(),
        name: data[i][1].toString().trim(),
        description: data[i][2].toString().trim(),
        body: data[i][3].toString()
      });
    }
  }
  return templates;
}
