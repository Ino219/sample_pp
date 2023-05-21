#pragma once

namespace samplepp {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// MyForm �̊T�v
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: �����ɃR���X�g���N�^�[ �R�[�h��ǉ����܂�
			//
		}

	protected:
		/// <summary>
		/// �g�p���̃��\�[�X�����ׂăN���[���A�b�v���܂��B
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
		/// �K�v�ȃf�U�C�i�[�ϐ��ł��B
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// �f�U�C�i�[ �T�|�[�g�ɕK�v�ȃ��\�b�h�ł��B���̃��\�b�h�̓��e��
		/// �R�[�h �G�f�B�^�[�ŕύX���Ȃ��ł��������B
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
		//�}�`���̍\����
		ref struct Ashapes {
			//�\���̖�
			String^ name;
			//�e�L�X�g���e
			String^ textVal;
			//�e�L�X�g�t�H���g��
			String^ fontName;
			//�e�L�X�g�T�C�Y
			int textSize;
			//�}�`�^�C�v(�e�L�X�g�A�^�C�g���A�摜�Ȃ�)
			Microsoft::Office::Core::MsoShapeType t;
			//�}�`�̍���
			int height;
			//�}�`�̕�
			int width;
			//�}�`��X���W
			int x;
			//�}�`��Y���W
			int y;
			//�e�L�X�g�^�C�v���ǂ���
			bool text;
			//�摜�^�C�v���ǂ���
			bool picture;
			//�摜�̏ꍇ�̓t�@�C���p�X
			String^ picturePath;
		};
		//�}�`���X�g
		Generic::List<Ashapes^>^ shapesList = gcnew Generic::List<Ashapes^>;
		//�e�}�`�\���̂��߂̃s�N�`���[�{�b�N�X
		PictureBox^ pic;
		//�e�L�X�g�\���̂��߂̃��x��
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
