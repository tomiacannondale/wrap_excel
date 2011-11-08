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
end
