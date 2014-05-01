#!/usr/local/bin/ruby
# -*- coding: cp932 -*-

require 'win32ole'
require 'delegate'
require 'singleton'
require 'util'

module VISIO
  module VisioConsts
  end

  class VisioDevice
    include Singleton
    include MessageUtil

    def initialize
      @visio = WIN32OLE.new('Visio.Application')
      @visio.Visible = true

      WIN32OLE.const_load @visio, VisioConsts
      @visio.Documents.OpenEx("f:/KD/work/管図面作製プログラム/PartsStencil.vss", 2)
      # @stendoc = @visio.Documents.OpenEx("g:/KD/work/管図面作製プログラム/PartsStencil2.vss", 2)

      @macdoc = @visio.Documents.Item(1)
      @doc = @visio.Documents.Add("")
      @curpage = "0"

      # デフォルトのページは立上げ時のページ
      @pagetable = {@curpage => @doc.Pages(1)}

      select_page(@curpage)
      @scale = 100   # デフォルトでは1/100の縮尺
      set_page
      @fstsel = true
    end

    def device
      @device
    end

    # ページサイズの設定
    def set_page_internal(pgd)
      pgd.PageSheet.Cells("DrawingScaleType").ResultIU = 4
      pgd.PageSheet.Cells("DrawingSizeType").ResultIU = 1
      pgd.PageSheet.Cells("PageScale").Formula = "1cm"
      pgd.PageSheet.Cells("DrawingScale").Formula = "#{@scale}cm"
      pgd.PageSheet.Cells("PageWidth").Formula = "10m"
      pgd.PageSheet.Cells("PageHeight").Formula = "10m"
    end

    def set_page
      set_page_internal(@device)
    end

    def set_scale(scale)
      @scale = scale
      set_page
    end

    def select_page(name)
      if @pagetable[name] == nil then
        npage = @doc.Pages.Add
        @pagetable[name] = npage
        set_page_internal(npage)
      end

      @device = @pagetable[name]
    end

    def select(obj)
      @visio.ActiveWindow.Select obj, 2
      @fstsel = false
    end

    def deselect
      @fstsel = true
      @visio.ActiveWindow.DeselectAll
    end

    def group
      @visio.ActiveWindow.Selection.Group
    end

    def set_line_attribute(line, width, style, color)
      line.Cells("LineWeight").Formula = "#{width} pt"
      line.Cells("LinePattern").Formula = style.to_s
      line.Cells("LineColor").Formula = color.to_s
    end

    def set_rect_attribute(rect, color)
      rect.Cells("FillForegnd").Formula = color.to_s
    end

    def draw_line(x1, y1, x2, y2)
      @device.DrawLine(x1, y1, x2, y2)
    end

    def draw_back(x1, y1, width, height)
      sha = @device.DrawRectangle(x1 - width, y1 - height, x1 + width, y1 + height)
      sha.Cells("LinePattern").Formula = "0"
      sha.Cells("FillPattern").Formula = "1"
      sha.Cells("FillBkgnd").Formula = "1"

      sha
    end

    def draw_rectangle(x1, y1, width, height)
      width /= 2.0
      height /= 2.0
      sha = @device.DrawRectangle(x1 - width, y1 - height, x1 + width, y1 + height)
      sha.Cells("LinePattern").Formula = "0"
      sha.Cells("FillPattern").Formula = "1"
      sha.Cells("FillBkgnd").Formula = "0"
      sha.Cells("FillForegnd").Formula = "0"
      sha
    end

    def draw_oval(x1, y1, x2, y2)
      sha = @device.DrawOval(x1, y1, x2, y2)
      sha
    end

    def draw_text(x1, y1, text, angle, fontsize = 8)
      width, height = text_size(text)
      xoff = ((width * 6 + 8) * @scale) / 100.0
#      xoff *= 1.5
      yoff = (6 * @scale + 4) / 100.0
#      yoff *= 1.5
      sha = @device.DrawRectangle(x1 - xoff, y1 - yoff, x1 + xoff, y1 + yoff)
      sha.text = text
      sha.Cells("Char.Size").Formula = "#{fontsize}pt"
      sha.Cells("LinePattern").Formula = "0"
      sha.Cells("FillPattern").Formula = "0"
      sha.Cells("Angle").Formula = angle.to_s + " deg"

      sha
    end

    def get_stencil(name)
      st = @stendoc.Masters.item(name)

      st
    end

    def drop(obj, x, y)
      #               print obj.name, " ", x, "  ", y, "\n"
      sh = @device.Drop(obj, x, y)
      ow = sh.Cells("Width").ResultIU
      sh.Cells("Width").ResultIU =  ow / 4
      oh = sh.Cells("Height").ResultIU
      sh.Cells("Height").ResultIU =  oh / 4

      sh
    end

    def draw_bezier(xy, degree, flag)
      # @device.DrawBezier(xy, degree, flag)
      @macdoc.draw_bezier(@device, *xy)
    end

    def draw_triangle(x1, y1, x2, y2, x3, y3)
      @macdoc.draw_triangle(@device, x1, y1, x2, y2, x3, y3)
    end

    def draw_polyline(xy)
      @macdoc.draw_polyline(xy)
    end

# Visioの引き出し線を使ったバージョン
=begin
    def draw_label(orgx, orgy, labx, laby, mess, width, height, wang)
      hiki = get_stencil("hikidasi")
      sh = drop(hiki, orgx, orgy)
      sh.Cells("BeginX").ResultIU = orgx
      sh.Cells("BeginY").ResultIU = orgy
      sh.Cells("EndX").ResultIU = labx
      sh.Cells("EndY").ResultIU = laby

      sh.Cells("Height").ResultIU = width
      sh.text = mess
      sh.Cells("Char.Size").Formula = "8pt"
    end
=end


    def draw_label(orgx, orgy, labx, laby, mess, width, height, wang)
      swidth = width
      sheight = height
      tlabx = labx - swidth / 2
      tlaby = laby

      if wang == 0 then
        tlabx = tlabx + swidth
      end

      lin = draw_line(orgx, orgy, labx, laby)
      set_line_attribute(lin, 0.5, 1, 0)
      draw_text(tlabx, tlaby, mess, 0, 8)
    end

    def finish
    end

    def save_as(fn)
      @doc.SaveAs(fn)
    end

    def quit
      @visio.Quit
    end

    def close
      @doc.Close
      @doc = @visio.Documents.Add("")
      @curpage = "0"

      # デフォルトのページは立上げ時のページ
      @pagetable = {@curpage => @doc.Pages(1)}

      select_page(@curpage)
      @scale = 100   # デフォルトでは1/100の縮尺
      set_page
      @fstsel = true
    end
  end
end # module VISIO
