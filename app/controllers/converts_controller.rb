class ConvertsController < ApplicationController
  before_action :set_convert, only: [:show, :edit, :update, :destroy, :download]

  # GET /converts or /converts.json
  def index
    @convert = Convert.new
  end

  # GET /converts/1 or /converts/1.json
  def show
    ext = File.extname(@convert.file_identifier)  #拡張子を取得
    # メッセージセット      
    case ext.downcase
    when ".csv" then
      @msgtitle = ' 注意事項'
      @msg1 = '・エクセルで設定されていた書式（日付や数値等）や数式はすべて解除され、セルの値が文字列で出力されます'
      @msg2 = '・複数シートが存在する場合、左端のシートが変換対象です'
      @msg3 = '・文字コードはUTF-8です'
      @msg4 = '・出力項目はすべて ””（ダブルクォーテーション）で囲んで出力します'
      @msg5 = '・このページの有効期限は5分間です'
      @msg6 = ''
      @msg7 = ''
      @msg8 = ''
      @msg9 = ''
    when ".xlsx" then
      @msgtitle = ' 注意事項'
      @msg1 = '・データはすべて文字列形式で出力されます'
      @msg2 = '・このページの有効期限は5分間です'
      @msg3 = ''
      @msg4 = ''
      @msg5 = ''
      @msg6 = ''
      @msg7 = ''
      @msg8 = ''
      @msg9 = ''
    else
      @msgtitle = '注意事項'
      @msg1 = '・このページの有効期限は5分間です'
      @msg2 = ''
      @msg3 = ''
      @msg4 = ''
      @msg5 = ''
      @msg6 = ''
      @msg7 = ''
      @msg8 = ''
      @msg9 = ''
    end    
  end

  # GET /converts/new
  def new
    @convert = Convert.new
  end

  # GET /converts/1/edit
  def edit
  end

  # POST /converts or /converts.json
  def create    
    require 'date'  #変換後ファイル名設定用
    require 'csv'
    require 'nkf'
    require 'rubyXL'
    require 'rubyXL/convenience_methods'    # フォントの変更機能を有効にするため
    require 'mini_magick'
    # require 'RMagick'

    @convert = Convert.new(convert_params) 

    if @convert.save
      # 選択ファイルのセット
      file = @convert.file

      # ファイル選択エラーチェック
      if file.blank?
        flash[:alert] = "ファイルを選択して下さい"
        redirect_to action: :index
        return
      end

      # 初期設定
      array = []  # 配列を宣言
      ext = File.extname(@convert.file_identifier)  #拡張子を取得
      modified_name = File.basename(@convert.file_identifier, ".*")     # 選択ファイルの元ファイル名取得（拡張子無し）

      #拡張子による処理分岐
      case ext.downcase   # 取得した拡張子を小文字変換して比較（.downcase）
      when ".csv" then

        # アップロードファイルのサイズチェック
        if File.size(@convert.file.current_path) > 5000000
          flash[:alert] = "5MBを超えるファイルは処理できません"
          redirect_to action: :index
          return
        end

        # 指定ファイルを3バイト分読み込んで、文字コードを判定
        enc_check = File.read(file.path,3)
        enc = NKF.guess(enc_check)
        
        case enc.name
        when "UTF-8" then
            # 配列に格納(UTF-8)
            CSV.foreach(file.path, quote_char: "\x00", col_sep: ",", encoding: "BOM:UTF-8") do |line|
                array.push(line) # arrayに格納する
            end
        when "EUC-JP" then
            # 配列に格納(EUC-JP)
            CSV.foreach(file.path, quote_char: "\x00", col_sep: ",", encoding: "BOM:UTF-8") do |line|
                array.push(line) # arrayに格納する
            end
        when "US-ASCII" then
            # 配列に格納(US-ASCII)
            CSV.foreach(file.path, quote_char: "\x00", col_sep: ",", encoding: "CP932:UTF-8") do |line|
                array.push(line) # arrayに格納する
            end
        when "Shift_JIS" then
            # 配列に格納(S-JIS)
            CSV.foreach(file.path, quote_char: "\x00", col_sep: ",", encoding: "CP932:UTF-8") do |line|
                array.push(line) # arrayに格納する
            end
        when "UTF-16" then
          # 配列に格納(Unicode)
          CSV.foreach(file.path, quote_char: "\x00", col_sep: ",", encoding: "UTF-16:UTF-8") do |line|
              array.push(line) # arrayに格納する
          end
        else
          flash[:alert] = "変換できない文字コードが含まれています"
          redirect_to action: :index
          return
        end
          
        # convert
        # 新規でブックを開く
        workbook = RubyXL::Workbook.new                
        # フォントのデフォルト指定
        workbook.fonts[0].set_name('ＭＳ ゴシック')
        # 一番左のシートを指定
        worksheet = workbook[0]
        # 編集する
        # i：配列の数（＝エクセルの行）
        # x：配列の要素数（＝エクセルの列）
        array.each_with_index do |arrays, i|
          arrays.each_with_index do |var, x|
            if arrays[x]
              arrays[x].gsub!("\"","") 
              cell = worksheet.add_cell i, x, arrays[x]
            else
              cell = worksheet.add_cell i, x, ""
            end                        
            x += 1
          end 
        end

        # 編集したファイル名のセット
        outputpath = File::dirname(file.path) #指定ファイルからディレクトリのみ取得
        modified_name += ".xlsx"
        outputpath = outputpath + "/" + modified_name
        workbook.write(outputpath)

        # 編集したファイル名にDB登録内容を変換
        @convert.update_columns(file:modified_name)
          
      when ".tsv" then

        # アップロードファイルのサイズチェック
        if File.size(@convert.file.current_path) > 5000000
          flash[:alert] = "5MBを超えるファイルは処理できません"
          redirect_to action: :index
          return
        end

        # 指定ファイルを3バイト分読み込んで、文字コードを判定
        enc_check = File.read(file.path,2)
        enc = NKF.guess(enc_check)

        case enc.name
        when "UTF-8" then
          # 配列に格納(UTF-8)
          CSV.foreach(file.path, quote_char: "\x00", col_sep: "\t", encoding: "BOM:UTF-8") do |line|
              array.push(line) # arrayに格納する
          end
        when "EUC-JP" then
          # 配列に格納(EUC-JP)
          CSV.foreach(file.path, quote_char: "\x00", col_sep: "\t", encoding: "BOM:UTF-8") do |line|
              array.push(line) # arrayに格納する
          end
        when "US-ASCII" then
          # 配列に格納(US-ASCII)
          CSV.foreach(file.path, quote_char: "\x00", col_sep: "\t", encoding: "CP932:UTF-8") do |line|
              array.push(line) # arrayに格納する
          end
        when "Shift_JIS" then
          # 配列に格納(S-JIS)
          CSV.foreach(file.path, quote_char: "\x00", col_sep: "\t", encoding: "CP932:UTF-8") do |line|
              array.push(line) # arrayに格納する
          end
        when "UTF-16" then
          # 配列に格納(Unicode)
          CSV.foreach(file.path, quote_char: "\x00", col_sep: "\t", encoding: "UTF-16:UTF-8") do |line|
              array.push(line) # arrayに格納する
          end
        else
          flash[:alert] = "変換できない文字コードが含まれています"
          redirect_to action: :index
          return
        end
        
        # convert
        # 新規でブックを開く
        workbook = RubyXL::Workbook.new                
        # フォントのデフォルト指定
        workbook.fonts[0].set_name('ＭＳ ゴシック')
        # 一番左のシートを指定
        worksheet = workbook[0]
        # 編集する
        # i：配列の数（＝エクセルの行）
        # x：配列の要素数（＝エクセルの列）
        array.each_with_index do |arrays, i|
          arrays.each_with_index do |var, x|
            if arrays[x]
              arrays[x].gsub!("\"","") 
              cell = worksheet.add_cell i, x, arrays[x]
            else
              cell = worksheet.add_cell i, x, ""
            end                        
            x += 1
          end 
        end         
        
        # 編集したファイル名のセット
        outputpath = File::dirname(file.path) #指定ファイルからディレクトリのみ取得
        modified_name += ".xlsx"
        outputpath = outputpath + "/" + modified_name
        workbook.write(outputpath)

        # 編集したファイル名にDB登録内容を変換
        @convert.update_columns(file:modified_name)
          
      when ".xlsx" then 
        # convert
        # BOMの付与
        bom = "\uFEFF"
        # ファイルを開く
        workbook = RubyXL::Parser.parse file.path
        # 一番左のシートを指定
        worksheet = workbook[0]
        # 編集する
        # Encoding::CP65001はutf-8を表現
        csv = CSV.generate(bom, row_sep: "\r\n", force_quotes: true, encoding:Encoding::CP65001) { |generate_data|
          # rubyXLでは、セル空白はnilと判定されてエラーとなるため、一括配列セットは廃止
          # worksheet.each_with_index do |arrays, i|
          #  set_data = worksheet[i].cells.map(&:value)  # 行単位での一括配列セット
          #  generate_data << set_data
          # end
          # 1行1セルずつ読み込む
          worksheet.each_with_index do |row, i|
            arrays = []
            row.cells.each do |cell|
              if(cell==nil)
                arrays.push("")
              else
                # 文字列に強制変換
                cell.set_number_format('@')
                col = cell.value
                arrays.push(col)
                # arrays.push(cell.value)
              end
            end
            generate_data << arrays
          end
        }
        
        # 編集したファイルをセット
        outputpath = File::dirname(file.path) #指定ファイルからディレクトリのみ取得
        modified_name += ".csv"        
        outputpath = outputpath + "/" + modified_name
        file = File.open(outputpath,"wb")
          file.print csv
        file.close

        # 編集したファイル名にDB登録内容を変換
        @convert.update_columns(file:modified_name)
        
      when ".heic" then 
        
        image = MiniMagick::Image.open(file.path)
        image.path
        image.format 'jpeg'       

        # RMagick使用時
        # image = Magick::Image.read(file.path).first

        # 編集したファイルをセット
        outputpath = File::dirname(file.path) #指定ファイルからディレクトリのみ取得
        modified_name += ".jpg"        
        outputpath = outputpath + "/" + modified_name

        image.write(outputpath)

        # 編集したファイル名にDB登録内容を変換
        @convert.update_columns(file:modified_name)

      else
        flash[:alert] = "変換できるファイル形式は xlsx csv tsv heic のみです"
        redirect_to action: :index
        return
      end      
    end
      redirect_to convert_url(@convert)
  end

  def download
    send_data(@convert.file.read,:filename=>@convert.file_identifier)
  end

  # PATCH/PUT /converts/1 or /converts/1.json
  def update
  end

  # DELETE /converts/1 or /converts/1.json
  def destroy
    @convert.destroy

    # flash[:notice] = ""
    redirect_to action: :index
    return
  end

  private
    # Use callbacks to share common setup or constraints between actions.
    def set_convert
      @convert = Convert.find(params[:id])
    end

    # Only allow a list of trusted parameters through.
    def convert_params
      # params.require(:convert).permit(:file)
      params.fetch(:convert, {}).permit(:file)
    end

end
