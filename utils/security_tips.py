import streamlit as st

def show_privacy_tips():
    """
    プライバシー保護のためのTipsを表示
    """
    with st.expander("🔒 プライバシー保護のためのTips", expanded=False):
        st.markdown("""
        ### ファイル処理の履歴を残さない方法

        #### 1. ブラウザの履歴とダウンロード履歴をクリア
        - **Chrome**: Ctrl+Shift+Delete → 「ダウンロード履歴」と「閲覧履歴」をチェック
        - **Firefox**: Ctrl+Shift+Delete → 「ダウンロード履歴」と「履歴」をチェック  
        - **Safari**: Cmd+Option+E → 「すべての履歴」を選択
        - **Edge**: Ctrl+Shift+Delete → 「ダウンロード履歴」と「閲覧履歴」をチェック

        #### 2. プライベートブラウジングモードを使用
        - **Chrome**: Ctrl+Shift+N (シークレットモード)
        - **Firefox**: Ctrl+Shift+P (プライベートウィンドウ)
        - **Safari**: Cmd+Shift+N (プライベートブラウザ)
        - **Edge**: Ctrl+Shift+N (InPrivateブラウザ)

        #### 3. 一時ファイルの自動削除
        - システムは一時ファイルを自動的に安全削除します
        - ファイルは暗号化された状態で処理されます
        - 処理完了後は完全に削除されます

        #### 4. 追加のセキュリティ対策
        - VPNを使用してIPアドレスを隠蔽
        - 共有PCでの使用は避ける
        - 処理後はブラウザを完全に閉じる
        - 機密データは暗号化されたストレージに保存

        #### 5. ファイル名の匿名化
        - アップロード・ダウンロード時にファイル名を匿名化
        - 元のファイル名は表示されません
        - ランダムな識別子を使用
        """)

def show_security_warning():
    """
    セキュリティ警告を表示
    """
    st.error("""
    ⚠️ **セキュリティ警告**
    
    機密情報を含むファイルを取り扱う際は、必ず以下を実行してください：
    1. プライベートブラウジングモードを使用
    2. 処理後はブラウザの履歴を完全にクリア
    3. 共有PCでの使用は避ける
    4. 必要に応じてVPNを使用
    """)

def show_cleanup_instructions():
    """
    クリーンアップ手順を表示
    """
    st.info("""
    🧹 **処理完了後のクリーンアップ**
    
    プライバシー保護のため、以下の手順でクリーンアップを実行してください：
    1. ブラウザの履歴をクリア
    2. ダウンロードフォルダから不要なファイルを削除
    3. システムの一時ファイルを削除
    4. 可能であればブラウザを再起動
    """)

def get_browser_cleanup_script():
    """
    ブラウザクリーンアップ用のJavaScriptコード
    """
    return """
    // ブラウザキャッシュとストレージをクリア
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.getRegistrations().then(function(registrations) {
            for(let registration of registrations) {
                registration.unregister();
            }
        });
    }
    
    // ローカルストレージをクリア
    localStorage.clear();
    sessionStorage.clear();
    
    // IndexedDBをクリア
    if ('indexedDB' in window) {
        indexedDB.databases().then(databases => {
            databases.forEach(db => {
                indexedDB.deleteDatabase(db.name);
            });
        });
    }
    
    console.log('ブラウザデータをクリーンアップしました');
    """