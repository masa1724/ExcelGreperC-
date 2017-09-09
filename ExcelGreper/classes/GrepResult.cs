
namespace ExcelGreper.classes
{
    /// <summary>
    /// Grep結果の1行保持
    /// </summary>
    class GrepResult
    {
        /// <summary>
        /// キーワードが見つかったオブジェクトの種類(セル、シェイプ等)
        /// </summary>
        public Types Type { get; private set; }

        /// <summary>
        /// キーワードが見つかったファイルのパス
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// キーワードが見つかったセルのアドレス
        /// シェイプの場合は、シェイプが配置されている左上のセルのアドレス
        /// </summary>
        public string CellAddress { get; private set; }

        /// <summary>
        /// キーワードが見つかったシェイプの名前
        /// this.TypeがShape(シェイプ)の場合のみ当プロパティに値が設定されます.
        /// </summary>
        public string ShapeName { get; private set; }

        /// <summary>
        /// キーワードが見つかったセル、シェイプのテキスト
        /// </summary>
        public string Text { get; private set; }

        /// <summary>
        /// オブジェクト種類の列挙型
        /// </summary>
        public enum Types
        {
            Cell, Shape
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="type">オブジェクトの種類</param>
        /// <param name="filePath">ファイルのパス</param>
        /// <param name="cellAddress">セルのアドレス</param>
        /// <param name="shapeName">シェイプの名前</param>
        /// <param name="text">セル、シェイプのテキスト</param>
        public GrepResult(Types type, string filePath, string cellAddress, string shapeName, string text)
        {
            this.Type = type;
            this.FilePath = filePath;
            this.CellAddress = cellAddress;
            this.ShapeName = shapeName;
            this.Text = text;
        }
    }
}
