/**
 * AuthService.gs - 権限管理サービス
 * ユーザー権限と共有設定を管理
 */

class AuthService {
  constructor() {
    this.spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  }

  /**
   * 現在のユーザー情報と権限を取得
   */
  getCurrentUser() {
    const user = Session.getActiveUser();
    const email = user.getEmail();
    
    // 共有設定を取得
    const sharedUsers = this.getSharedUsers();
    const userPermission = sharedUsers.find(u => u.email === email);
    
    return {
      email: email,
      name: this.getUserName(email),
      permission: userPermission ? userPermission.permission : 'view',
      isOwner: this.isOwner(email)
    };
  }

  /**
   * 共有ユーザーリストを取得
   */
  getSharedUsers() {
    try {
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      let permSheet = ss.getSheetByName('_permissions');
      
      if (!permSheet) {
        permSheet = this.createPermissionsSheet(ss);
      }
      
      const data = permSheet.getDataRange().getValues();
      if (data.length <= 1) return [];
      
      return data.slice(1).map(row => ({
        email: row[0],
        name: row[1],
        permission: row[2],
        addedDate: row[3],
        addedBy: row[4]
      })).filter(user => user.email);
    } catch (error) {
      Logger.log('Error getting shared users: ' + error.toString());
      return [];
    }
  }

  /**
   * ユーザーを追加
   */
  addSharedUser(email, permission = 'view') {
    if (!this.validateEmail(email)) {
      return { success: false, error: '無効なメールアドレスです' };
    }
    
    try {
      const currentUser = this.getCurrentUser();
      if (!this.canManagePermissions(currentUser)) {
        return { success: false, error: '権限を管理する権限がありません' };
      }
      
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      let permSheet = ss.getSheetByName('_permissions');
      
      if (!permSheet) {
        permSheet = this.createPermissionsSheet(ss);
      }
      
      // 既存チェック
      const existingUsers = this.getSharedUsers();
      if (existingUsers.some(u => u.email === email)) {
        return { success: false, error: 'このユーザーは既に追加されています' };
      }
      
      // 新規追加
      const newRow = [
        email,
        this.getUserName(email),
        permission,
        new Date(),
        currentUser.email
      ];
      
      permSheet.appendRow(newRow);
      
      // 実際のスプレッドシート権限も付与
      this.grantSpreadsheetAccess(email, permission);
      
      return { 
        success: true, 
        user: {
          email: email,
          name: this.getUserName(email),
          permission: permission
        }
      };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  /**
   * ユーザー権限を更新
   */
  updateUserPermission(email, newPermission) {
    try {
      const currentUser = this.getCurrentUser();
      if (!this.canManagePermissions(currentUser)) {
        return { success: false, error: '権限を管理する権限がありません' };
      }
      
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      const permSheet = ss.getSheetByName('_permissions');
      
      if (!permSheet) {
        return { success: false, error: '権限シートが見つかりません' };
      }
      
      const data = permSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === email) {
          permSheet.getRange(i + 1, 3).setValue(newPermission);
          
          // スプレッドシート権限も更新
          this.grantSpreadsheetAccess(email, newPermission);
          
          return { success: true };
        }
      }
      
      return { success: false, error: 'ユーザーが見つかりません' };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  /**
   * ユーザーを削除
   */
  removeSharedUser(email) {
    try {
      const currentUser = this.getCurrentUser();
      if (!this.canManagePermissions(currentUser)) {
        return { success: false, error: '権限を管理する権限がありません' };
      }
      
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      const permSheet = ss.getSheetByName('_permissions');
      
      if (!permSheet) {
        return { success: false, error: '権限シートが見つかりません' };
      }
      
      const data = permSheet.getDataRange().getValues();
      
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][0] === email) {
          permSheet.deleteRow(i + 1);
          
          // スプレッドシート権限も削除
          this.revokeSpreadsheetAccess(email);
          
          return { success: true };
        }
      }
      
      return { success: false, error: 'ユーザーが見つかりません' };
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  /**
   * 権限管理シートを作成
   */
  createPermissionsSheet(ss) {
    const sheet = ss.insertSheet('_permissions');
    
    // ヘッダー
    const headers = ['Email', 'Name', 'Permission', 'Added Date', 'Added By'];
    sheet.getRange('A1:E1').setValues([headers]);
    sheet.getRange('A1:E1').setFontWeight('bold');
    sheet.getRange('A1:E1').setBackground('#e5e7eb');
    
    // オーナーを追加
    const owner = Session.getActiveUser().getEmail();
    sheet.appendRow([owner, this.getUserName(owner), 'owner', new Date(), 'System']);
    
    // シートを保護
    const protection = sheet.protect();
    protection.setDescription('System sheet - Do not edit');
    protection.setWarningOnly(true);
    
    // シートを非表示に
    sheet.hideSheet();
    
    return sheet;
  }

  /**
   * スプレッドシートへのアクセス権限を付与
   */
  grantSpreadsheetAccess(email, permission) {
    try {
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      
      if (permission === 'edit' || permission === 'owner') {
        ss.addEditor(email);
      } else {
        ss.addViewer(email);
      }
    } catch (error) {
      Logger.log('Error granting access: ' + error.toString());
    }
  }

  /**
   * スプレッドシートへのアクセス権限を削除
   */
  revokeSpreadsheetAccess(email) {
    try {
      const ss = SpreadsheetApp.openById(this.spreadsheetId);
      ss.removeEditor(email);
      ss.removeViewer(email);
    } catch (error) {
      Logger.log('Error revoking access: ' + error.toString());
    }
  }

  /**
   * ユーザー名を取得（簡易実装）
   */
  getUserName(email) {
    // 実際の実装では、Directory APIやPeople APIを使用
    return email.split('@')[0];
  }

  /**
   * メールアドレスの検証
   */
  validateEmail(email) {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
  }

  /**
   * オーナーかどうか判定
   */
  isOwner(email) {
    const sharedUsers = this.getSharedUsers();
    const user = sharedUsers.find(u => u.email === email);
    return user && user.permission === 'owner';
  }

  /**
   * 権限管理が可能か判定
   */
  canManagePermissions(user) {
    return user.permission === 'owner' || user.permission === 'edit';
  }

  /**
   * 共有リンクを生成
   */
  generateShareLink() {
    const baseUrl = ScriptApp.getService().getUrl();
    const token = Utilities.getUuid();
    
    // トークンを保存（24時間有効）
    PropertiesService.getScriptProperties().setProperty(`share_token_${token}`, JSON.stringify({
      created: new Date().getTime(),
      permission: 'view'
    }));
    
    return `${baseUrl}?token=${token}`;
  }

  /**
   * 共有トークンを検証
   */
  validateShareToken(token) {
    const tokenData = PropertiesService.getScriptProperties().getProperty(`share_token_${token}`);
    
    if (!tokenData) {
      return { valid: false };
    }
    
    const data = JSON.parse(tokenData);
    const now = new Date().getTime();
    const dayInMs = 24 * 60 * 60 * 1000;
    
    if (now - data.created > dayInMs) {
      PropertiesService.getScriptProperties().deleteProperty(`share_token_${token}`);
      return { valid: false };
    }
    
    return { valid: true, permission: data.permission };
  }
}
