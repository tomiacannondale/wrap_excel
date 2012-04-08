# -*- coding: utf-8 -*-
require File.join(File.dirname(__FILE__), './spec_helper')

describe WrapExcel::Cell do
  before do
    @dir = create_tmpdir
  end

  after do
    rm_tmp(@dir)
  end

  context "open simple.xls" do
    before do
      @book = WrapExcel::Book.open(@dir + '/simple.xls')
      @sheet = @book[1]
      @cell = @sheet[0, 0]
    end

    after do
      @book.close
    end

    describe "#value" do
      it "get cell's value" do
        @cell.value.should eq 'simple'
      end
    end

    describe "#value=" do
      it "change cell data to 'fooooo'" do
        @cell.value = 'fooooo'
        @cell.value.should eq 'fooooo'
      end
    end

    describe "#method_missing" do
      context "unknown method" do
        it { expect { @cell.hogehogefoo }.to raise_error }
      end
    end

  end

  context "open merge_cells.xls" do
    before do
      @book = WrapExcel::Book.open(@dir + '/merge_cells.xls')
      @sheet = @book[0]
    end

    after do
      @book.close
    end

    it "merged cell get same value" do
      @sheet[0, 0].value.should be_nil
      @sheet[1, 0].value.should eq 'first merged'
    end

    it "set merged cell" do
      @sheet[1, 0].value = "set merge cell"
      @sheet[1, 0].value.should eq "set merge cell"
      @sheet[1, 1].value.should eq "set merge cell"
    end
  end
end
