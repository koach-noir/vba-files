import os
import re

def find_eum_macros(directory):
    eum_macros = []
    # 正規表現パターン - Sub または Function で始まり、名前の後に _EUM が付くものを検索
    pattern = r'(Sub|Function)\s+([^\s\(]+)_EUM'
    
    # ディレクトリを再帰的に検索
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.bas'):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='shift_jis', errors='ignore') as f:
                        content = f.read()
                        # 正規表現でマクロ名を検索
                        matches = re.findall(pattern, content)
                        for match in matches:
                            macro_name = match[1] + "_EUM"
                            eum_macros.append(macro_name)
                except Exception as e:
                    print(f"エラー: {file_path} を読み込めませんでした: {e}")
    
    return eum_macros

def main():
    # スクリプトが置かれているディレクトリを取得
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # マクロを検索
    macros = find_eum_macros(current_dir)
    
    # 重複を削除して並べ替え
    macros = sorted(list(set(macros)))
    
    # 結果をファイルに書き込み
    output_file = os.path.join(current_dir, "EUMNameList.txt")
    with open(output_file, 'w', encoding='shift_jis') as f:
        for macro in macros:
            f.write(f"{macro}\n")
    
    print(f"{len(macros)}個のEUMマクロが見つかりました")
    print(f"結果は {output_file} に保存されました")

if __name__ == "__main__":
    main()
