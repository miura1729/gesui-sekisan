#!/usr/local/bin/ruby -d -Ks
# -*- mode: ruby; tab-width: 2;  -*-
require 'win32ole'
require 'delegate'
require 'singleton'

module ExcelConsts
end

class TheExcel
  include Singleton

  def initialize
    @excel = WIN32OLE.new("Excel.Application")
    @excel.Visible = TRUE
    WIN32OLE.const_load @excel, ExcelConsts
  end

  def method_missing(m, *args, &block)
#    @excel.__send(m, *args, &block)
    @excel.send(m, *args, &block)
  end

  def TheExcel.quit
    TheExcel.instance.quit
  end

  def TheExcel.visible=(switch)
#    TheExcel.instance.Visible = switch
    TheExcel.instance.Visible = switch
  end
end

class CellArray
  def initialize(cell)
    @cell = cell
  end

  def cell_raw(n)
    @cell.cells(1, n + 1)
  end

  def [](n)
    @cell.cells(1, n + 1).value
  end

  def []=(n, v)
    @cell.cells(1, n + 1).formula = v
  end
end

class ExcelScan
  def initialize(fname, sname = nil, visible = true)
    @excel = TheExcel.instance
    @excel.Visible = visible
    @book = @excel.Workbooks.open fname
    @sheet = nil
    @sheet_list = {}
    if sname != nil then
      @book.Worksheets.each do |ws|
        if ws.name == sname then
          @sheet = ws
        end
        @sheet_list[ws.name] = ws
      end
    else
      @book.Worksheets.each do |ws|
        if @sheet == nil then
          @sheet = ws
        end
        @sheet_list[ws.name] = ws
      end
    end

    @cur_row = 1
  end

  def set_sheet(sname)
    @sheet = @sheet_list[sname]
    if @sheet == nil
      raise "No such sheet #{sname}"
    end
  end

  def add_sheet(sname, orgsheet)
    @sheet_list[orgsheet].Copy(@sheet_list[orgsheet])
    newsheet = nil
    @book.Worksheets.each do |ws|
      if !@sheet_list[ws.name] then
        newsheet = ws
      end
    end
#   newsheet = @book.Worksheets.Add
#   newsheet.Paste
    newsheet.name = sname
    @sheet_list[sname] = newsheet
    @sheet = newsheet
  end

  def set_row(n)
    @cur_row = n
  end

  # セルの書式設定を行う
  #  rowはカレント、colを指定する
  def set_attr(col, bsym, color)
    sym2index = {:Left => ExcelConsts::XlEdgeLeft, 
                :Top  => ExcelConsts::XlEdgeTop,
                :Bottom => ExcelConsts::XlEdgeBottom,
                :Right => ExcelConsts::XlEdgeRight}
    curcell = @sheet.cells(@cur_row, col + 1)
#   sym2index.values.each do |i|
#     b = curcell.Borders(i)
#     b.LineStyle = ExcelConsts::XlLineStyleNone
#   end
    bsym.each do |sym|
      b = curcell.Borders(sym2index[sym])
      b.ColorIndex = ExcelConsts::XlAutomatic
      b.LineStyle = ExcelConsts::XlContinuous
      b.Weight = ExcelConsts::XlThin
    end
    if color then
      curcell.Interior.ColorIndex = color
    end
  end

  def insert(n = nil)
    if n then
      @cur_row = n
    end

    @sheet.Rows(@cur_row).Insert
    nl = CellArray.new(@sheet.Rows(@cur_row))
    @cur_row = @cur_row + 1

    nl
  end

  def rows
    CellArray.new(@sheet.Rows(@cur_row))
  end

  def reset
    @cur_row = 1
  end

  def next
    @cur_row = @cur_row + 1
    rows
  end

  def each
    lcnt = @cur_row
    ura = @sheet.UsedRange
    if ura != nil
      rend = ura.Rows.Count
    else
      rend = 0
    end
    while lcnt <= rend
      yield CellArray.new(@sheet.rows(lcnt))
      lcnt = lcnt + 1
    end
  end

  def each_sheet
    @sheet_list.each do |sn, sh|
      lcnt = @cur_row
      rend = sh.UsedRange.Rows.Count
      while lcnt <= rend
        yield CellArray.new(sh.rows(lcnt))
        lcnt = lcnt + 1
      end
    end
  end

  def close
    @book.Close
  end

  def close!
    TheExcel.instance['DisplayAlerts'] = false
    @book.Close
  end

  def saveas(fn)
    @book.SaveAs(fn)
  end

  def set_page(prop)
    ps = @sheet.PageSetup
    ps.LeftHeader = prop['LeftHeader'] if defined? prop['LeftHeader']
    ps.RightHeader = prop['RightHeader'] if defined? prop['RightHeader']
    ps.CenterHeader = prop['CenterHeader'] if defined? prop['CenterHeader']

    ps.LeftFooter = prop['LeftFooter'] if defined? prop['LeftFooter']
    ps.RightFooter = prop['RightFooter'] if defined? prop['RightFooter']
    ps.CenterFooter = prop['CenterFooter'] if defined? prop['CenterFooter']
  end
end       
