/**
 * DataService.gs - データ操作専用サービス
 * 全てのデータ操作をこのファイルに集約
 */

class DataService {
  constructor() {
    this.spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    this.cache = CacheService.getScriptCache();
  }

  /**
   * 複数の組織計画を管理
   */
  getAllPlans() {
    const cacheKey = 'all_plans';
    const cached = this.cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    try {
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      const sheets = ss.getSheets();
      const plans = {};

      sheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.startsWith('_')) return; // システムシートはスキップ
        
        // シートのメタデータを取得
        const metaRange = sheet.getRange('A1:B3');
        const metaData = metaRange.getValues();
        
        const planInfo = {
          id: sheetName,
          name: metaData[0][1] || sheetName,
          period: metaData[1][1] || '',
          memo: metaData[2][1] || '',
          nodes: this.getNodesFromSheet(sheet)
        };
        
        plans[sheetName] = planInfo;
      });

      // キャッシュに保存（5分間）
      this.cache.put(cacheKey, JSON.stringify(plans), 300);
      return plans;
    } catch (error) {
      throw new Error(`データ取得エラー: ${error.message}`);
    }
  }

  /**
   * シートからノードデータを取得
   */
  getNodesFromSheet(sheet) {
    const dataRange = sheet.getRange('A5:L' + sheet.getLastRow());
    const data = dataRange.getValues();
    const headers = data[0];
    
    return data.slice(1)
      .filter(row => row[0]) // IDが存在する行のみ
      .map(row => {
        const node = {};
        headers.forEach((header, index) => {
          switch(header) {
            case 'ID':
              node.id = row[index]?.toString() || '';
              break;
            case '名前':
              node.name = row[index] || '';
              break;
            case '役職':
              node.position = row[index] || '';
              break;
            case 'グレード':
              node.grade = row[index] || '';
              break;
            case 'レベル':
              node.level = parseInt(row[index]) || 0;
              break;
            case '上司ID':
              node.parentId = row[index]?.toString() || null;
              break;
            case 'タイプ':
              node.type = row[index] || 'existing';
              break;
            case '雇用形態':
              node.employment = row[index] || '正社員';
              break;
            case 'X座標':
              node.x = parseFloat(row[index]) || 0;
              break;
            case 'Y座標':
              node.y = parseFloat(row[index]) || 0;
              break;
            case '整列済み':
              node.isArranged = row[index] === true;
              break;
          }
        });
        return node;
      });
  }

  /**
   * 計画を保存（新規作成or更新）
   */
  savePlan(planId, planData) {
    try {
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      let sheet = ss.getSheetByName(planId);
      
      // 新規作成の場合
      if (!sheet) {
        // テンプレートシートをコピー
        const template = ss.getSheetByName('_template') || this.createTemplate(ss);
        sheet = template.copyTo(ss);
        sheet.setName(planId);
      }
      
      // メタデータを更新
      sheet.getRange('B1').setValue(planData.name);
      sheet.getRange('B2').setValue(planData.period);
      sheet.getRange('B3').setValue(planData.memo || '');
      
      // ノードデータを更新
      this.updateNodesInSheet(sheet, planData.nodes);
      
      // キャッシュをクリア
      this.cache.remove('all_plans');
      
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  /**
   * シートのノードデータを更新
   */
  updateNodesInSheet(sheet, nodes) {
    // 既存データをクリア（ヘッダー以外）
    const lastRow = sheet.getLastRow();
    if (lastRow > 5) {
      sheet.getRange(6, 1, lastRow - 5, 12).clearContent();
    }
    
    if (nodes && nodes.length > 0) {
      const rows = nodes.map(node => [
        node.id || '',
        node.name || '',
        node.position || '',
        node.grade || '',
        node.level || 0,
        node.parentId || '',
        node.type || 'existing',
        node.employment || '正社員',
        node.x || 0,
        node.y || 0,
        node.isArranged || false,
        new Date() // 最終更新日時
      ]);
      
      sheet.getRange(6, 1, rows.length, 12).setValues(rows);
    }
  }

  /**
   * テンプレートシートを作成
   */
  createTemplate(ss) {
    const template = ss.insertSheet('_template');
    
    // メタデータ領域
    template.getRange('A1:A3').setValues([['計画名:'], ['実施時期:'], ['メモ:']]);
    template.getRange('A1:B3').setBorder(true, true, true, true, true, true);
    
    // ヘッダー行
    const headers = [
      'ID', '名前', '役職', 'グレード', 'レベル', '上司ID', 
      'タイプ', '雇用形態', 'X座標', 'Y座標', '整列済み', '更新日時'
    ];
    template.getRange('A5:L5').setValues([headers]);
    template.getRange('A5:L5').setFontWeight('bold');
    template.getRange('A5:L5').setBackground('#f3f4f6');
    
    // シートを非表示に
    template.hideSheet();
    
    return template;
  }

  /**
   * 計画を削除
   */
  deletePlan(planId) {
    if (planId === 'current') {
      return { success: false, error: '現在の組織は削除できません' };
    }
    
    try {
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      const sheet = ss.getSheetByName(planId);
      
      if (sheet) {
        ss.deleteSheet(sheet);
        this.cache.remove('all_plans');
        return { success: true };
      }
      
      return { success: false, error: 'シートが見つかりません' };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  /**
   * ノードの一括更新（効率化）
   */
  batchUpdateNodes(planId, updates) {
    try {
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      const sheet = ss.getSheetByName(planId);
      
      if (!sheet) {
        return { success: false, error: 'シートが見つかりません' };
      }
      
      // 現在のデータを取得
      const currentPlan = this.getAllPlans()[planId];
      const nodes = currentPlan.nodes;
      
      // 更新を適用
      updates.forEach(update => {
        const nodeIndex = nodes.findIndex(n => n.id === update.id);
        if (nodeIndex !== -1) {
          nodes[nodeIndex] = { ...nodes[nodeIndex], ...update.data };
        }
      });
      
      // シートに反映
      this.updateNodesInSheet(sheet, nodes);
      this.cache.remove('all_plans');
      
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }
}
