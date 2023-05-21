#pragma once

namespace samplepp {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// MyForm の概要
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: ここにコンストラクター コードを追加します
			//
		}

	protected:
		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::PictureBox^  pictureBox1;
	protected:

	private:
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		void InitializeComponent(void)
		{
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->SuspendLayout();
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::SystemColors::ActiveCaptionText;
			this->pictureBox1->Location = System::Drawing::Point(153, 25);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(305, 295);
			this->pictureBox1->TabIndex = 0;
			this->pictureBox1->TabStop = false;
			// 
			// MyForm
			// 
			this->AllowDrop = true;
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 12);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(911, 431);
			this->Controls->Add(this->pictureBox1);
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			this->DragDrop += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::MyForm_DragDrop);
			this->DragEnter += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::MyForm_DragEnter);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->ResumeLayout(false);

		}
#pragma endregion
		//図形情報の構造体
		ref struct Ashapes {
			//構造体名
			String^ name;
			//テキスト内容
			String^ textVal;
			//テキストフォント名
			String^ fontName;
			//テキストサイズ
			int textSize;
			//図形タイプ(テキスト、タイトル、画像など)
			Microsoft::Office::Core::MsoShapeType t;
			//図形の高さ
			int height;
			//図形の幅
			int width;
			//図形のX座標
			int x;
			//図形のY座標
			int y;
			//テキストタイプかどうか
			bool text;
			//画像タイプかどうか
			bool picture;
			//画像の場合はファイルパス
			String^ picturePath;
		};
		//図形リスト
		Generic::List<Ashapes^>^ shapesList = gcnew Generic::List<Ashapes^>;
		//各図形表示のためのピクチャーボックス
		PictureBox^ pic;
		//テキスト表示のためのラベル
		Label^ tx;
		bool text_pickUp = false;
		int picX;
		int picY;
		Image^ im;
		
	private: System::Void MyForm_DragDrop(System::Object^  sender, System::Windows::Forms::DragEventArgs^  e);
	private: System::Void MyForm_DragEnter(System::Object^  sender, System::Windows::Forms::DragEventArgs^  e);
	private: System::Void picture_DragDrop(System::Object^  sender, System::Windows::Forms::DragEventArgs^  e);
	private: System::Void picture_DragEnter(System::Object^  sender, System::Windows::Forms::DragEventArgs^  e);
	private: System::Void picture_MouseDown(System::Object^  sender, System::Windows::Forms::MouseEventArgs^  e);
	private: System::Void MyForm_Load(System::Object^  sender, System::EventArgs^  e);
	private: System::Void text_MouseDown(System::Object^  sender, System::Windows::Forms::MouseEventArgs^  e);
	};
}
