/**
 * Code.gs - メインAPI
 * 各サービスを統合して、クライアントに提供
 */

// グローバルインスタンス
let dataService;
let authService;

/**
 * 初期化
 */
function init() {
  if (!dataService) dataService = new DataService();
  if (!authService) authService = new AuthService();
}

/**
 * Webアプリのエントリーポイント
 */
function doGet(e) {
  init();
  
  // 共有トークンの処理
  if (e.parameter.token) {
    const validation = authService.validateShareToken(e.parameter.token);
    if (!validation.valid) {
      return HtmlService.createHtmlOutput('無効なリンクです');
    }
  }
  
  // メインページを返す
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('組織デザインラボ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * HTMLファイルのインクルード
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 初期設定をセットアップ
 */
function setupInitialConfig() {
  const props = PropertiesService.getScriptProperties();
  
  // スプレッドシートIDが設定されていない場合
  if (!props.getProperty('SPREADSHEET_ID')) {
    // 新しいスプレッドシートを作成
    const ss = SpreadsheetApp.create('組織デザインラボ - データ');
    props.setProperty('SPREADSHEET_ID', ss.getId());
    
    // 初期シートをセットアップ
    const sheet = ss.getActiveSheet();
    sheet.setName('current');
    
    // テンプレートを作成
    init();
    dataService.createTemplate(ss);
    
    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl()
    };
  }
  
  return {
    success: true,
    spreadsheetId: props.getProperty('SPREADSHEET_ID')
  };
}

// ========== データ操作API ==========

/**
 * 全ての組織計画を取得
 */
function getAllPlans() {
  init();
  try {
    const plans = dataService.getAllPlans();
    const currentUser = authService.getCurrentUser();
    
    return {
      success: true,
      data: plans,
      user: currentUser
    };
  } catch (error) {
    Logger.log('Error in getAllPlans: ' + error.toString());
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 計画を保存
 */
function savePlan(planId, planData) {
  init();
  
  // 権限チェック
  const currentUser = authService.getCurrentUser();
  if (currentUser.permission === 'view') {
    return { success: false, error: '編集権限がありません' };
  }
  
  return dataService.savePlan(planId, planData);
}

/**
 * 計画を削除
 */
function deletePlan(planId) {
  init();
  
  // 権限チェック
  const currentUser = authService.getCurrentUser();
  if (currentUser.permission !== 'owner' && currentUser.permission !== 'edit') {
    return { success: false, error: '削除権限がありません' };
  }
  
  return dataService.deletePlan(planId);
}

/**
 * ノードの一括更新
 */
function batchUpdateNodes(planId, updates) {
  init();
  
  // 権限チェック
  const currentUser = authService.getCurrentUser();
  if (currentUser.permission === 'view') {
    return { success: false, error: '編集権限がありません' };
  }
  
  return dataService.batchUpdateNodes(planId, updates);
}

/**
 * CSV出力
 */
function exportToCSV(planId) {
  init();
  
  try {
    const plans = dataService.getAllPlans();
    const plan = plans[planId];
    
    if (!plan) {
      return { success: false, error: '計画が見つかりません' };
    }
    
    const headers = ['ID', '名前', '役職', 'グレード', 'レベル', '上司ID', 'タイプ', '雇用形態', 'X座標', 'Y座標'];
    const rows = [headers];
    
    plan.nodes.forEach(node => {
      rows.push([
        node.id,
        node.name,
        node.position,
        node.grade || '',
        node.level,
        node.parentId || '',
        node.type,
        node.employment,
        node.x,
        node.y
      ]);
    });
    
    const csvContent = rows.map(row => 
      row.map(cell => {
        if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"') || cell.includes('\n'))) {
          return '"' + cell.replace(/"/g, '""') + '"';
        }
        return cell;
      }).join(',')
    ).join('\n');
    
    return {
      success: true,
      data: csvContent,
      filename: `組織図_${plan.name}_${new Date().toISOString().split('T')[0]}.csv`
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ========== 権限管理API ==========

/**
 * 共有ユーザーリストを取得
 */
function getSharedUsers() {
  init();
  
  try {
    const users = authService.getSharedUsers();
    return { success: true, data: users };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * ユーザーを追加
 */
function addSharedUser(email, permission) {
  init();
  return authService.addSharedUser(email, permission);
}

/**
 * ユーザー権限を更新
 */
function updateUserPermission(email, permission) {
  init();
  return authService.updateUserPermission(email, permission);
}

/**
 * ユーザーを削除
 */
function removeSharedUser(email) {
  init();
  return authService.removeSharedUser(email);
}

/**
 * 共有リンクを生成
 */
function generateShareLink() {
  init();
  
  try {
    const link = authService.generateShareLink();
    return { success: true, link: link };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ========== ユーティリティ ==========

/**
 * 開発用：サンプルデータを生成
 */
function generateSampleData() {
  init();
  
  const sampleNodes = [
    { id: '1', name: '田中 太郎', position: '代表取締役', grade: 'G8', level: 0, parentId: null, x: 500, y: 50, type: 'existing', employment: '正社員' },
    { id: '2', name: '佐藤 花子', position: '営業本部長', grade: 'G7', level: 1, parentId: '1', x: 200, y: 180, type: 'existing', employment: '正社員' },
    { id: '3', name: '山田 次郎', position: '開発本部長', grade: 'G7', level: 1, parentId: '1', x: 500, y: 180, type: 'existing', employment: '正社員' },
    { id: '4', name: '鈴木 美香', position: '管理本部長', grade: 'G7', level: 1, parentId: '1', x: 800, y: 180, type: 'existing', employment: '正社員' },
  ];
  
  const currentPlan = {
    name: '現在の組織',
    period: new Date().toISOString().slice(0, 7),
    memo: '現在の組織構成',
    nodes: sampleNodes
  };
  
  return dataService.savePlan('current', currentPlan);
}
