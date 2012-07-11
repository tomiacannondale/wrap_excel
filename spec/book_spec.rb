# -*- coding: utf-8 -*-

require File.join(File.dirname(__FILE__), './spec_helper')

describe WrapExcel::Book do
  before do
    @dir = create_tmpdir
    @simple_file = @dir + '/simple.xls'
  end

  after do
    rm_tmp(@dir)
  end

  describe ".open" do
    context "exist file" do
      it "simple file with default" do
        expect {
          book = WrapExcel::Book.open(@simple_file)
          book.close
        }.to_not raise_error
      end

      it "simple file with writable" do
        expect {
          book = WrapExcel::Book.open(@simple_file, :read_only => false)
          book.close
        }.to_not raise_error
      end

      it "simple file with visible = true" do
        expect {
          book = WrapExcel::Book.open(@simple_file, :visible => true)
          book.close
        }.to_not raise_error
      end

      context "with block" do
        it 'block parameter should be instance of WrapExcel::Book' do
          WrapExcel::Book.open(@simple_file) do |book|
            book.should be_is_a WrapExcel::Book
          end
        end
      end
    end
  end

  describe 'access sheet' do
    before do
      @book = WrapExcel::Book.open(@simple_file)
    end

    after do
      @book.close
    end

    it 'with sheet name' do
      @book['Sheet1'].should be_kind_of WrapExcel::Sheet
    end

    it 'with integer' do
      @book[0].should be_kind_of WrapExcel::Sheet
    end

    it 'with block' do
      @book.each do |sheet|
        sheet.should be_kind_of WrapExcel::Sheet
      end
    end

    context 'open with block' do
      it {
        WrapExcel::Book.open(@simple_file) do |book|
          book['Sheet1'].should be_is_a WrapExcel::Sheet
        end
      }
    end
  end

  describe "#add_sheet" do
    before do
      @book = WrapExcel::Book.open(@simple_file)
      @sheet = @book[0]
    end

    after do
      @book.close
    end

    context "only first argument" do
      it "should add worksheet" do
        expect { @book.add_sheet @sheet }.to change{ @book.book.Worksheets.Count }.from(3).to(4)
      end

      it "should return copyed sheet" do
        sheet = @book.add_sheet @sheet
        copyed_sheet = @book.book.Worksheets.Item(@book.book.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end
    end

    context "with first argument" do
      context "with second argument is {:as => 'copyed_name'}" do
        it "copyed sheet name should be 'copyed_name'" do
          @book.add_sheet(@sheet, :as => 'copyed_name').name.should eq 'copyed_name'
        end
      end

      context "with second argument is {:before => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(@sheet, :before => @sheet).name.should eq @book[0].name
        end
      end

      context "with second argument is {:after => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(@sheet, :after => @sheet).name.should eq @book[1].name
        end
      end

      context "with second argument is {:before => @book[2], :after => @sheet}" do
        it "should arguments in the first is given priority" do
          @book.add_sheet(@sheet, :before => @book[2], :after => @sheet).name.should eq @book[2].name
        end
      end

    end

    context "without first argument" do
      context "second argument is {:as => 'new sheet'}" do
        it "should return new sheet" do
          @book.add_sheet(:as => 'new sheet').name.should eq 'new sheet'
        end
      end

      context "second argument is {:before => @sheet}" do
        it "should add the first sheet" do
          @book.add_sheet(:before => @sheet).name.should eq @book[0].name
        end
      end

      context "second argument is {:after => @sheet}" do
        it "should add the second sheet" do
          @book.add_sheet(:after => @sheet).name.should eq @book[1].name
        end
      end

    end

    context "without argument" do
      it "should add empty sheet" do
        expect { @book.add_sheet }.to change{ @book.book.Worksheets.Count }.from(3).to(4)
      end

      it "shoule return copyed sheet" do
        sheet = @book.add_sheet
        copyed_sheet = @book.book.Worksheets.Item(@book.book.Worksheets.Count)
        sheet.name.should eq copyed_sheet.name
      end
    end
  end

  describe ".save" do
    context "when open with read only" do
      before do
        @book = WrapExcel::Book.open(@simple_file)
      end

      it {
        expect {
          @book.save
        }.to raise_error(IOError,
                     "Not opened for writing(open with :read_only option)")
      }
    end

  end

end
