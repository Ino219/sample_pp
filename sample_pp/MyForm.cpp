#include "MyForm.h"

using namespace samplepp;
using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;
using namespace System::Collections::Generic;


[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew MyForm());
	return 0;
}

System::Void samplepp::MyForm::MyForm_DragDrop(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//フォームへのドラッグドロップイベント
	array<String^>^ file = (array<String^>^)e->Data->GetData(DataFormats::FileDrop, false);
	//拡張子の取得
	String^	extension = System::IO::Path::GetExtension(file[0]);
	//フルパス取得
	String^ path = System::IO::Path::GetFullPath(file[0]);

	Microsoft::Office::Interop::PowerPoint::Application^ app_ = nullptr;
	List<Microsoft::Office::Interop::PowerPoint::Shape^>^ shapeList = gcnew List<Microsoft::Office::Interop::PowerPoint::Shape^>;

	if (extension == ".pptx") {
		//パワーポイントファイルの場合、処理を開始
		app_ = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
		Microsoft::Office::Interop::PowerPoint::Presentations^ presen = app_->Presentations;
		Microsoft::Office::Interop::PowerPoint::Presentation^ presense = presen->Open(
			path,
			MsoTriState::msoFalse,
			MsoTriState::msoFalse,
			MsoTriState::msoFalse
		);
		//図形名
		String^ name;
		//テキストフレーム
		Microsoft::Office::Interop::PowerPoint::TextFrame^ text;
		//テキスト内容
		String^ text2;
		//画像パス
		String^ picPath;
		//フォント名
		String^ font;
		//フォントサイズ
		int TextSize;
		//パネル内項目数
		int itemCount = 0;
		//テキストかどうか
		bool textCheck;
		//画像かどうか
		bool pictureCheck;
		//表かどうか
		bool tableCheck;
		//図形スタイル
		MsoShapeStyleIndex style;
		//図形タイプ
		MsoShapeType type;
		int height;
		int width;
		int x;
		int y;

		
		//picturebox表示用
		Bitmap^ b;

		//タイトルがあれば、以下のやり方でタイトル情報を取得できる
		//String^ titleText = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Text;
		//String^ titlefont = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Font->Name;
		//String^ titlefont2 = presense->Slides[1]->Shapes->Title->AlternativeText;
		//int title_size = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Font->Size;

		

		//presense->Slides[1]->Shapes->Title->Export("C:\\Users\\chach\\Desktop\\sample_title.jpg", PpShapeFormat::ppShapeFormatJPG, width, height, PpExportMode::ppClipRelativeToSlide);

		//図形を一つずつ取得し、処理
		for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in presense->Slides[1]->Shapes)
		{
			//スライド表示用のキャンバス幅
			int slideWidth = presense->PageSetup->SlideWidth;
			//スライド表示用のキャンバス高さ
			int slideHeight = presense->PageSetup->SlideHeight;
			pictureBox1->Width = slideWidth;
			pictureBox1->Height = slideHeight;
			//描画用ビットマップ
			b = gcnew Bitmap(pictureBox1->Width, pictureBox1->Height);
			//真偽値変数初期化
			pictureCheck = false;
			textCheck = false;
			//図形の名前を取得
			name = var->Name;
			//図形のタイプを取得
			style = var->ShapeStyle;
			type = var->Type;
			//図形の高さ
			height = var->Height;
			//図形の幅
			width = var->Width;
			//図形の位置
			x = var->Left;
			y = var->Top;

			
			//テキストタイプの処理
			if (var->HasTextFrame==MsoTriState::msoTrue) {
				textCheck = true;
				//テキストが含まれている場合
				text = var->TextFrame;
				//テキスト取得
				text2 = text->TextRange->Text;
				//テキストサイズ
				TextSize = text->TextRange->Font->Size;
				//テキストフォント名
				font = var->AlternativeText;
				//var->Export("C:\\Users\\chach\\Desktop\\sample_text"+text+".jpg", PpShapeFormat::ppShapeFormatJPG, width, height, PpExportMode::ppScaleXY);
			}

			if (var->HasTable == MsoTriState::msoTrue) {
				pictureCheck = true;
				//テーブルを保有していた場合、一旦は画像として保持
				var->Export("C:\\Users\\chach\\Desktop\\sample_hyou.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleXY);
				picPath = "C:\\Users\\chach\\Desktop\\sample_hyou.bmp";
			}
			//埋め込みデータの場合
			if (type==MsoShapeType::msoEmbeddedOLEObject) {
				pictureCheck = true;
				var->Export("C:\\Users\\chach\\Desktop\\sample_umekomi.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleXY);
				picPath = "C:\\Users\\chach\\Desktop\\sample_umekomi.bmp";
			}
			//画像ファイルの場合
			if (type == MsoShapeType::msoPicture) {
				pictureCheck = true;
				var->Export("C:\\Users\\chach\\Desktop\\sample_pic.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleXY);
				picPath = "C:\\Users\\chach\\Desktop\\sample_pic.bmp";
			}
			//図形構造体に代入
			Ashapes^ temp = gcnew Ashapes;
			temp->name = name;
			temp->t = type;
			temp->height = height;
			temp->width = width;
			temp->x = x;
			temp->y = y;
			temp->text = textCheck;
			temp->textVal = text2;
			temp->picture = pictureCheck;
			temp->picturePath = picPath;
				
			
			//図形情報を取得
			shapeList->Add(var);
			shapesList->Add(temp);
		}
		int count = 0;
		//取得した図形情報をパネルとpictureBoxに表示処理していく
		for each (Ashapes^ var in shapesList)
		{
			//画像の場合
			if (var->picture) {
				//パネル表示用のピクチャーボックス
				pic = gcnew PictureBox;
				//コントロールに追加
				this->Controls->Add(pic);
				//ピクチャーボックスの座標
				pic->Top = 25 * itemCount;
				//高さ
				pic->Height = 25;
				//画像パスからイメージ作成
				System::Drawing::Image^ img= System::Drawing::Image::FromFile(var->picturePath);
				//イメージのはめ込み
				pic->Image = img;
				//表示用のピクチャーボックスのイメージを空のビットマップに設定
				pictureBox1->Image = b;
				//グラフィックを設定
				Graphics^ gr = Graphics::FromImage(pictureBox1->Image);
				//画像表示
				gr->DrawImage(img, var->x, var->y, var->width, var->height);
				//描画更新
				pictureBox1->Refresh();
				count++;
				pic->MouseDown += gcnew System::Windows::Forms::MouseEventHandler(this, &MyForm::picture_MouseDown);
				
			}
			else {
				//ラベル作成
				tx = gcnew Label;
				//コントロールに追加
				this->Controls->Add(tx);
				tx->Top = 25 * itemCount;
				tx->Text = var->textVal;
				//ピクチャーボックス作成
				//空のビットマップに文字描画
				//フォントなど指定
				tx->MouseDown += gcnew System::Windows::Forms::MouseEventHandler(this, &MyForm::text_MouseDown);
			}
			itemCount++;
		}
		//パワーポイントのインスタンス処理
		presense->Close();
		app_->Quit();

	}
	return System::Void();
}

System::Void samplepp::MyForm::MyForm_DragEnter(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//フォームへのファイルドラッグドロップイベント
	if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		e->Effect = DragDropEffects::All;
	}
	else {
		e->Effect = DragDropEffects::None;
	}
	return System::Void();
}

System::Void samplepp::MyForm::picture_DragDrop(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//パネルからのドラッグドロップイベント
	if (e->Data->GetDataPresent(DataFormats::Bitmap)) {
		System::Drawing::Graphics^ tegr = pictureBox1->CreateGraphics();
		Image^ img = (Image^)e->Data->GetData(DataFormats::Bitmap);
		int xp = pictureBox1->MousePosition.X * img->Width / pictureBox1->Width;
		int yp = pictureBox1->MousePosition.Y * img->Height / pictureBox1->Height;
		tegr->DrawImage((Image^)e->Data->GetData(DataFormats::Bitmap), xp, yp);
		//pictureBox1->Image = (Image^)e->Data->GetData(DataFormats::Bitmap);

	}
	else if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		//データオブジェクトの受け取り
		array<String^>^ temp=(array<String^>^)e->Data->GetData(DataFormats::FileDrop);
		
	}
	return System::Void();
}

System::Void samplepp::MyForm::picture_DragEnter(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		e->Effect = DragDropEffects::All;
	}
	else {
		e->Effect = DragDropEffects::None;
	}
	return System::Void();
}

System::Void samplepp::MyForm::picture_MouseDown(System::Object ^ sender, System::Windows::Forms::MouseEventArgs ^ e)
{
	//マウスクリックしたピクチャーボックスを取得
	pic = (PictureBox^)sender;
	//ドラッグドロップイベントを設定
	pic->DoDragDrop(pic->Image, DragDropEffects::All);
	return System::Void();
}

System::Void samplepp::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	//ロード時、ドロップイベントを許可
	pictureBox1->AllowDrop = true;
	//イベント駆動を設定
	pictureBox1->DragDrop += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::picture_DragDrop);
	pictureBox1->DragEnter += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::picture_DragDrop);
	return System::Void();
}


System::Void samplepp::MyForm::text_MouseDown(System::Object ^ sender, System::Windows::Forms::MouseEventArgs ^ e)
{
	//パネルからラベルを選択した場合の処理
	text_pickUp = true;
	//必要なデータの配列を作成
	array<String^>^ data = gcnew array<String^>(3);
	//どのラベルが選択されたか
	tx = (Label^)sender;
	//ラベルのテキストを取得
	String^ lbtext = tx->Text;
	//テキストと一致する図形情報を検索し、テキスト内容、フォント、サイズを取得し、配列に格納
	for each (Ashapes^ var in shapesList)
	{
		if (var->text&&(lbtext == var->textVal)) {
			data[0] = var->textVal;
			data[1] = var->fontName;
			data[2] = var->textSize.ToString();
		}
	}
	//データオブジェクトとしてまとめ、ファイルドロップ形式で紐づける
	DataObject^ dobj = gcnew DataObject(DataFormats::FileDrop, data);
	//ドラッグドロップイベントでデータオブジェクトを引っ張る
	tx->DoDragDrop(dobj, DragDropEffects::All);
	return System::Void();
}

