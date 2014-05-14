#!/usr/local/bin/ruby -Ks -w
# -*- coding: cp932 -*-
# -*- mode: ruby; tab-width: 2;  -*-

require 'visio'
require 'excelwriter'
require 'config'
$pipe150 = 0

class List
  def initialize(car, cdr)
    @car = car
    @cdr = cdr
  end

  def [](n)
    c = self
    while n > 0 do
      n = n - 1
      c = c.cdr
      unless c then
        raise "Unexpected nil #{self}"
      end
    end

    c.car
  end

  def to_s
    inspect
  end

  def inspect
    "(" + inspect_aux + ")"
  end

  def inspect_aux
    ca = "nil"
    cd = "nil"
    if @car then
      ca = @car.inspect
    end

    if @cdr then
      if @cdr.kind_of?(List) then
        cd = @cdr.inspect_aux
      elsif @cdr != nil then
        cd = @cdr.inspect
      else
        return ca
      end
    end

    ca + " " + cd
  end

  attr_accessor :car
  attr_accessor :cdr
end

class Reader
  def initialize(stream)
    @stream = stream
  end

  def skip_space
    while /\s/ =~ (ch = @stream.read(1)) do
      if ch == nil then
        return nil
      end
    end
    return ch
  end

  def get_next_token
    ch = skip_space
    if ch == ';' then
      @stream.gets
      ch = skip_space
    end

    if ch == '"' then
      token = ''
      while ch = @stream.read(1) do
        if ch == '"' then
          break
        end

        token <<= ch
      end

      return token
    end

    token = ch
    while /\S/ =~ (ch = @stream.read(1)) do
      if ch == nil then
        break
      end

      if token == '(' then
        @stream.ungetc(ch[0])
        break
      end

      if ch == ')' then
        @stream.ungetc(ch[0])
        break
      end

      token <<= ch
    end

    if /^-?(([0-9]+)|([0-9]*\.[0-9]+))$/ =~ token then
      eval(token)
    elsif /^\".*\"$/ =~ token then
      eval(token)
    else
      token
    end
  end

  def read
    cl = nil
    rt = nil

    while tk = get_next_token do
      case tk
      when '('
        if cl == nil then
          cl = List.new(read, nil)
          rt = cl
        else
          cl.cdr = List.new(read, nil)
          cl = cl.cdr
        end

      when ')'
        return rt

      else
        if cl == nil then
          cl = List.new(tk, nil)
          rt = cl
        else
          cl.cdr = List.new(tk, nil)
          cl = cl.cdr
        end
      end
    end

    rt
  end
end

class Parser
  def parse(tree)
    case tree.car

    when 'K'
      KoukyouMasu.new.parse(tree)

    when 'KD'
      DropKoukyouMasu.new.parse(tree)

    when 'N'
      NoneKoukyouMasu.new.parse(tree)

    else
      raise "公共マスの定義を書きます K (直桝), KD (ドロップ桝) " +
        tree.car.to_s
    end
  end
end

class Geometry
  def matle2screen(n)
    # Visioはinch単位
    (n / 0.0254)
  end
end

class GeoPoint<Geometry
  def inside?(min, max)
    if (self.x < min.x) or (self.y < min.y) then
      return false
    end

    if (self.x > max.x) or (self.y > max.y) then
      return false
    end

    true
  end

  def initialize(x, y)
    @x = x
    @y = y
  end

  def x
    @x
  end

  def y
    @y
  end

  def scalex
    matle2screen(@x)
  end

  def scaley
    matle2screen(@y)
  end

  def roll(deg, org)
    rad = (deg * Math::PI) / 180.0
    tx = (@x - org.x) * Math.cos(rad) - (@y - org.y) * Math.sin(rad)
    ty = (@x - org.x) * Math.sin(rad) + (@y - org.y) * Math.cos(rad)
    GeoPoint.new(tx + org.x, ty + org.y)
  end

  def move!(step, angle)
    rad = (angle * Math::PI) / 180.0
    @x = @x + step * Math.cos(rad)
    @y = @y + step * Math.sin(rad)
  end

  def move(step, angle)
    rad = (angle * Math::PI) / 180.0
    x = @x + step * Math.cos(rad)
    y = @y + step * Math.sin(rad)

    GeoPoint.new(x, y)
  end

  def +(point)
    GeoPoint.new(@x + point.x, @y + point.y)
  end

  def -(point)
    GeoPoint.new(@x - point.x, @y - point.y)
  end

  def distance(point)
    dx = @x - point.x
    dy = @y - point.y
    Math.sqrt(dx * dx + dy * dy)
  end

  def midPoint(point)
    GeoPoint.new((@x + point.x) / 2, (@y + point.y) / 2)
  end

  def join_to_a(*join)
    res = [scaley, scalex]
    join.each do |n|
      res.push n.scaley
      res.push n.scalex
    end

    res.reverse
  end

  def to_s
    "(#{@x}, #{@y})"
  end

  ORIGIN = GeoPoint.new(0, 0)
end

class RangeSet
  def initialize
    @ranges = []
  end

  def add(range)
    min = range[0]
    max = range[1]
    _add(min, max)
    @ranges = @ranges.select {|r| r}
  end

  def _delete_subset(i, max)
    while i < @ranges.size and @ranges[i].last <= max do
      @ranges[i] = nil
      i = i + 1
    end

    if i < @ranges.size and @ranges[i].first <= max then
      newmax = @ranges[i].last
      @ranges[i] = nil
      newmax
    else
      max
    end
  end

  def _delete_subset2(i, min)
    while 0 <= i and @ranges[i].first >= min do
      @ranges[i] = nil
      i = i - 1
    end

    if 0 <= i and @ranges[i].last >= min then
      newmin = @ranges[i].first
      @ranges[i] = nil
      newmin
    else
      min
    end
  end

  def _add(min, max)
    @ranges.each_with_index do |r, i|
      if r.first <= min then
        if r.last >= min then
          if r.last < max
            newmax = _delete_subset(i + 1, max)
            @ranges[i] = Range.new(r.first, newmax)
          end

          return
        end
      else
        if r.last <= min then
          if r.last < max then
            newmax = _delete_subset(i + 1, max)
            @ranges[i] = Range.new(min, newmax)
          elsif r.first != min then
            newmin = _delete_subset2(i - 1, min)
            @ranges[i] = Range.new(newmin, r.last)
          end

          return
        end
      end
    end

    @ranges.push(Range.new(min, max))
    @ranges.sort! {|a, b| a.first <=> b.first}
  end

  def neg(min, max)
    cmin = min
    ocmin = cmin
    cmax = min
    res = []
    @ranges.each do |r|
      if r.last > max then
        if ocmin != cmax then
          res.push(Range.new(ocmin, cmax))
        end

        return res
      end

      if r.first > cmax then
        cmax = r.first
      end

      if r.last > cmin then
        cmin = r.last
      end

      if ocmin != cmax then
        res.push(Range.new(ocmin, cmax))
      end

      ocmin = cmin
    end

    if ocmin != max then
      res.push(Range.new(ocmin, max))
    end

    return res
  end

  def to_s
    @ranges
  end
end

class VectorBag
  def initialize
    @parm = []
    @temp = []
  end

  def parm_add(vec)
    @parm.push vec
  end

  def temp_add(vec)
    @temp.push vec
  end

  def temp_clear
    @temp = []
  end

  def temp_pop
    @temp.pop
  end

  def filter(point, len)
    res = []
    @parm.each do |v|
      if point.distance(v[0]) < len or
          point.distance(v[1]) < len then
        res.push v
      end
    end

    @temp.each do |v|
      if point.distance(v[0]) < len or
          point.distance(v[1]) < len then
        res.push v
      end
    end

    res
  end
end

class Vector
  def initialize(p0, p1)
    @points = [p0, p1]
  end

  def [](n)
    @points[n]
  end

  def angle_range(pt)
    res = []
    [0, 1].each do |i|
      tv = @points[i] - pt
      tvl = Math.sqrt(tv.x * tv.x + tv.y * tv.y)
      if tvl != 0 then
        res.push((Math.atan2(tv.y, tv.x) * 180 / Math::PI) % 360)
      end
    end

    if res.size == 1 then
      res = [res[0], res[0]]
    end

    if res.size == 0 then
      return []
    end

    if res[0] > res[1] then
      if res[0] - res[1] > 180 then
        res = [[0, res[1]], [res[0], 360]]
      else
        res = [[res[1], res[0]]]
      end
    else
      if res[1] - res[0] > 180 then
        res = [[0, res[0]], [res[1], 360]]
      else
        res = [[res[0], res[1]]]
      end
    end

    res
  end
end

module AngleResolver
  @@vector_bag = VectorBag.new

  def init_resolver
    collect_vector
    compute_label_angle(nil)
  end

  def label_angle
    if not(defined? @label_angle) or @label_angle == nil then
      @label_angle = alt_search_free_angle + @angle - 30
    end

    @label_angle
  end

  def add_vector(p1, p2)
    @@vector_bag.parm_add(Vector.new(p1, p2))
  end

  def collect_vector
    addvect
    if @enter then
      @@vector_bag.parm_add(Vector.new(@enter.pos, @pos))
    end

    if self.kind_of?(KoukyouMasu) then
      @@vector_bag.parm_add(Vector.new(@pos.move(0.01, @angle + 180), @pos))
    end

    @exit.each do |n|
      n[1].collect_vector
    end
  end

  # ヒントは上流側のラベルの角度。角度がそろっていたほうが格好いいので
  # 出来る限り角度をあわせる。ただし、
  #    ・ パイプや他のラベルとぶつかってしまう時
  #    ・ ラベル同士がくっついている時
  # はその限りではない。
  def compute_label_angle(hint)
    # あまり近い時はヒントがあるとラベルが重なってしまうので
    # 無視
=begin
    if @enter_length and @enter_length < 0.5 * (scale / 100) then
      hint = nil
    end
=end

    success = true
    @label_angle = nil
    newhint = hint

    while @label_angle == nil do
      if self.kind_of?(NotPrintParts) then
        @label_angle = 0  # 0はダミー。nilでなければなんでもいい
      else
        @label_angle = search_free_angle(hint, collect_disable_angle(scale))
      end

      if @label_angle == nil then
        # successがfalseになるということはこのテストが1度は通って
        # 次で失敗してwhile trueでもう1度戻ってきたって事なので
        # 1度目の成功時のpushをpopしてやる。
        if success == false then
          @@vector_bag.temp_pop
        end

        return nil
      else
        newhint = @label_angle
      end

      @exit.each do |n|
        if n[1].compute_label_angle(newhint) == nil then
          success = false
        end
      end

      if success then
        return @label_angle
      end
    end

    @label_angle
  end

  def collect_disable_angle(scale)
    va = @@vector_bag.filter(@pos, 5 * (scale / 100))
    # 12pt
    #			va = @@vector_bag.filter(@pos, 20 * (scale / 100))
    range_set = RangeSet.new
    va.each do |n|
      ang = n.angle_range(@pos)
      ang.each do |a|
        range_set.add(a)
      end
    end

    range_set
  end

  def search_free_angle(hint, range_set)
    sz = 0
    wr = nil
    free = range_set.neg(0, 360)
    free.each do |r|
      if hint and r === hint and r.first + 20 < hint and r.last - 20 > hint then
#        @@vector_bag.temp_add(Vector.new(@pos.move(5, hint), @pos))
        return hint
      end

      if sz < r.size
        sz = r.size
        wr = r
      end
    end

    if sz <= 0 then
      return nil
    end

    ang = (wr.last * 2 + wr.first) / 3
    @@vector_bag.temp_add(Vector.new(@pos.move(5, ang), @pos))
    ang
  end

  #  ベクターアルゴリズムが失敗した時の代替方法
  # 初期のアルゴリズム
  #
  def alt_search_free_angle
    angarr = [180, 360]
    @exit.each do |n|
      angarr.push n[0]
    end

    angarr.sort!

    if angarr == [0, 180, 360] then
      hm = @exit[0][1]
      if hm == nil then
        return 90
      else
        return hm.alt_search_free_angle
      end
    end

    prev = 0
    max = 0
    gang = 0
    angarr.each do |n|
      if max < (n - prev) then
        max = n - prev
        gang = (prev + n) / 2
      end

      prev = n
    end

    gang
  end
end

class Range
  def size
    last - first
  end
end

module DrawPrimitive
  def set_attr(device, sh)
    if enter_pipe_new == "OLD"
      device.set_line_attribute(sh, 0.72, 23, 0)
    else
      device.set_line_attribute(sh, 0.72, 1, 0)
    end
  end

  def draw_line(device, x1, y1, x2, y2)
    sh = device.draw_line(x1, y1, x2, y2)
    set_attr(device, sh)
    sh
  end

  def draw_oval(device, x1, y1, x2, y2)
    sh = device.draw_oval(x1, y1, x2, y2)
    set_attr(device, sh)
    sh
  end

  def draw_rectangle(device, x1, y1, x2, y2)
    sh = device.draw_rectangle(x1, y1, x2, y2)
    set_attr(device, sh)
    sh
  end
end

module DrawUtil
  include DrawPrimitive

  def draw_pipelen(device)
    if enter_pipe_new == "OLD" then
      return
    end

    midp = @pos.midPoint(@enter.pos)
    angle = @angle

    if angle >= 180 then
      angle = angle - 180
    end

    if @enter_length < 0.5 then
      # 12pt
      #		if @enter_length < 2 then
      midp.move!(0.2 * scale / 100, angle + 90)
    else
      midp.move!(0.1 * scale / 100, angle + 90)
    end

    if @enter_length != 0 then
      txtstr = "%.2f" % @enter_length
      if @enter_length > 3 and @koubai and @koubai != 0 then
        txtstr += " (%s)" % (((@koubai * 1000).to_i)/1000r)
      end
      device.draw_text(midp.scalex, midp.scaley, txtstr, angle)
    end
  end

  def vanilla_masu_draw(device)
    draw_pipelen(device)

    if enter_pipe_kind == "ヒューム管" or enter_pipe_kind == "鋼管" then
      draw_rectangle(device, @pos.scalex, @pos.scaley, 14, 14)
    end

    draw_oval(device, @pos.scalex - 5,
              @pos.scaley - 5,
              @pos.scalex + 5,
              @pos.scaley + 5)

    if $hdrawf then
      text = "No.%s H%2.2f GL %1.1f %s" % [@no, @level + @alevel, @alevel, name]
    else
      text = "No.%s GL %1.1f %s" % [@no, @alevel, name]
    end
    draw_label(device, text)
  end

  def draw_label(device, text)
    if enter_pipe_new == "OLD" then
      return
    end

    fangle = label_angle
    pos = @pos.move(0.5 * (scale / 100.0), fangle)
    tpos = pos.move(2 * (scale / 100.0), fangle)
    # 12pt
    #		pos = @pos.move(2 * (scale / 100.0), fangle)
    #		tpos = pos.move(8 * (scale / 100.0), fangle)
    lin = draw_line(device, pos.scalex,
                    pos.scaley,
                    tpos.scalex,
                    tpos.scaley)
    device.set_line_attribute(lin, 0.1, 1, 0)

    tpos2 = tpos.move(0.1 * (scale / 100.0), fangle + 90)
    txt = device.draw_text(tpos2.scalex,
                           tpos2.scaley,
                           text,
                           fangle)
    device.deselect
    device.select(txt)
    device.select(lin)
    device.group
    device.deselect
  end

  def drop_draw(device, size)
    x = @pos.scalex
    y = @pos.scaley
    rad = ((@angle + 180) * Math::PI) / 180.0
    x1 = x + size * Math.cos(rad)
    y1 = y + size * Math.sin(rad)

    sh = draw_line(device, x, y, x1, y1)
    sh.Cells("LinePattern").Formula = "1"
    sh.Cells("LineColor").Formula = "1"

    (-10..10).each do |i|
      r1 = ((@angle + i * 2 + 180) * Math::PI) / 180.0
      r2 = ((@angle + (i + 1) * 2 + 180) * Math::PI) / 180.0

      draw_line(device,x + size * Math.cos(r1),
                y + size * Math.sin(r1),
                x + size * Math.cos(r2),
                y + size * Math.sin(r2))
    end
  end

  def box_draw(device, xs, ys)
    xs2 = xs / 2
    ys2 = ys / 2
    x = @pos.scalex
    y = @pos.scaley
    rad = ((@angle + 180) * Math::PI) / 180.0
    x1 = x + ys2 * Math.cos(rad)
    y1 = y + ys2 * Math.sin(rad)

    sh = draw_rectangle(device, x1, y1, xs, ys)
    sh.Cells("Angle").Formula = ((@angle + 90) % 360).to_s  + " deg"
  end
end

module SekisanPipeMainParts
  @@used_pipe = Hash.new(0)
  def used_pipe
    @@used_pipe
  end

  def sekisan_pipe
    if @enter == nil then
      return
    end

    if enter_pipe_size >= 100 then
      sekisan_pipe_with_high
    else
      #				sekisan_pipe_without_high
      sekisan_pipe_with_high
    end
  end

  def sekisan_pipe_with_high
    if @enter_length == 0 then
      return
    end

    edlevel = @enter.level + @enter.alevel
    stlevel = @level + @alevel
    delta = edlevel - stlevel
    length = @enter_length

    prelevel = stlevel
    curlevel = (stlevel * 5 + 0.9999).to_i / 5.0
    while curlevel + 0.2 < edlevel do
      plen = ((length * (curlevel - prelevel)) / delta)
      plab = (curlevel).to_s + "■" + enter_pipe_size.to_s + "■" + enter_pipe_kind.to_s
      used_pipe[plab] += plen

      prelevel = curlevel
      curlevel = curlevel + 0.2
    end

    if delta == 0 then
      plen = length
    else
      plen = ((length * (edlevel - prelevel)) / delta)
    end

    plab = curlevel.to_s + "■" + enter_pipe_size.to_s + "■" + enter_pipe_kind.to_s
    used_pipe[plab] += plen
  end

  def sekisan_pipe_without_high
    plab = "■" + enter_pipe_size.to_s + "■" + enter_pipe_kind.to_s
    used_pipe[plab] += @enter_length
  end
end

module TankaCalcDB
  CONNECTION_STRING = 'DSN=kyuusui'

  def TankaCalcDB.get_cleanup(connect)
    proc {
      connect.Close
    }
  end

  def initialize
=begin
       @connect = WIN32OLE.new "ADODB.Connection"
       @connect.Open CONNECTION_STRING
       ObjectSpace.define_finalizer(self, TankaCalcDB.get_cleanup(@connect))
=end
  end

  def get_tanka(name, spec)
=begin
       begin
         query = "SELECT price FROM tanka WHERE name='#{name}' AND spec='#{spec}'"
         rs = @connect.Execute query
         if not rs.Eof then
           rt = rs.Fields.Item('price').Value
           rs.MoveNext
           if not rs.Eof then
             raise "価格リストが重複しています #{name} #{spec}"
           end
         else
           rt = nil
         end

       ensure
         rs.Close
       end
=end
    rt = nil
    rt
  end
end

class MitumorishoWirter
  def write(masu, pipe)
    wr_begin

    name = File.basename(ARGV[0], ".*") + "  様"
    wr_atesaki(name)

    wr_masu_begin
    masu.keys.sort {|a, b|
      si = ((a[1] == b[1]) ? 0 : 1)
      a[si] <=> b[si]
    }.each do |key|
      wr_masu(masu[key], key[0].split(/, /))
    end
    wr_masu_end

    wr_pipe_begin
    pipe.keys.sort.each do |key|
      wr_pipe(pipe[key], key)
    end
    wr_pipe_end

    wr_end
  end
end

class ExcelMitumorishoWriter<MitumorishoWirter
  include TankaCalcDB

  def wr_begin
    @scan = ExcelScan.new("f:\\KD\\work\\管図面作製プログラム\\mitumori.xls")
    @masu_num = 0
    @pipe_num = 0
  end

  def wr_atesaki(name)
    @scan.set_row(1)
    (@scan.rows)[0] = name
  end

  def wr_masu_begin
    @scan.set_row(15)
  end

  def wr_masu(num, item)
    sa = @scan.insert
    sa[0] = item[0]
    sa[1] = item[1]
    sa[2] = num
    sa[3] = "ヶ"

    if (pri = get_tanka(item[0], item[1])) then
      sa[5] = pri
      sa[6] = "=RC[-1] * RC[-4]"
    end

    @masu_num += 1
  end

  def wr_masu_end
    sa = @scan.insert
    sa[0] = "合計"
    sa[6] = "=SUM(R[-1]C:R[-#{@masu_num}]C)"
  end

  def wr_pipe_begin
    @scan.insert
    @scan.insert
  end

  def wr_pipe(len, spec)
    sa = @scan.insert
    name = "排水管"

    dpth, psize, pkind = spec.split(/■/)
    if dpth == "" then
      if pkind == "VU" then
        sspec = "φ#{psize}"
      else
        sspec = "φ#{psize}#{pkind}"
      end
    else
      if pkind == "VU" then
        sspec = "φ#{psize}×H#{dpth}～#{dpth.to_f + 0.2}"
      else
        sspec = "φ#{psize}#{pkind}×H#{dpth}～#{dpth.to_f + 0.2}"
      end
    end

    sa[0] = name
    sa[1] = sspec
    sa[2] = len
    sa[3] = "m"
    if (pri = get_tanka(name, sspec)) then
      sa[5] = pri
      sa[6] = "=RC[-1] * RC[-4]"
    end

    @pipe_num += 1
  end

  def wr_pipe_end
    sa = @scan.insert
    sa[0] = "合計"
    sa[6] = "=SUM(R[-1]C:R[-#{@pipe_num}]C)"
  end

  def wr_end
  end
end

class Parts
  include AngleResolver
  include DrawUtil

  @@scale = 100

  @@used_parts = Hash.new(0)

  def scale
    @@scale
  end

  def used_parts
    @@used_parts
  end

  def initialize
    @enter = nil
    @enter_length = nil
    @exit = []

    @pos = nil
    @angle = nil
    @level = nil
    @tori_level = 0
    @alevel = 0

    @koubai = 0.02    # 1/50
  end

  attr :pos
  attr :level
  attr :no
  attr_accessor :koubai

  def set_position(enter, length, angle)
    if $adjust_masulen then
      off = GeoPoint.new(length + 0.254 / 2, 0).roll(angle, GeoPoint::ORIGIN)
    else
      off = GeoPoint.new(length, 0).roll(angle, GeoPoint::ORIGIN)
    end

    @pos = enter.pos + off
    @enter = enter
    @enter_length = length
    @angle = angle
  end

  def set_pipe(size, kind, pnew, tlev, alevel)
    @pipe_size = size
    @pipe_kind = kind
    @pipe_new  = pnew
    @tori_level = tlev
    @alevel = alevel
  end

  attr :pipe_size
  attr :pipe_kind
  attr :pipe_new
  attr :tori_level
  attr :alevel

  def enter_pipe_size
    if @enter then
      @enter.pipe_size
    else
      pipe_size
    end
  end

  def enter_pipe_kind
    if @enter then
      @enter.pipe_kind
    else
      pipe_kind
    end
  end

  def enter_pipe_new
    if @enter then
      @enter.pipe_new
    else
      pipe_new
    end
  end

  # angle は出の方向,通り方向が0で後は度で指定
  def add_exit(exit, length, angle)
    @exit.push [angle, exit]

    exit.set_position(self, length, (@angle + angle) % 360)
    exit.set_pipe(@pipe_size, @pipe_kind, @pipe_new, @tori_level, @alevel)
    exit.calc_level
    if @level + @alevel < 0.1 then
      print "Recalc executed in #{self}\n"
      recalc_level(self, nil, 0)
      exit.calc_level
    end
  end

  #  文法
  #     公共マス, KDはドロップマス
  #    K (上流側深さ 縮尺 [角度]) 
  #      (放流反対側)
  #      (右側)
  #      (左側)
  # 
  #  インバートマス(チーズを兼ねる)
  #   I (右) (左) 真っ直ぐ
  #
  #  エルボ (0°にすると直マス)
  #   L 角度 ...
  #
  #  ドロップマス
  #   D 角度 ...
  #
  #  2方向ドロップマス
  #   DD 角度 (右) (左) 
  #    角度は真ん中を基準(普通は0)
  #
  #  トラップマス
  #   T (右) (左) 真っ直ぐ
  #
  #  マスの直後に文字列を書くとラベルの設定
  #  この場合は自動採番はされません。
  #  例  I "10-1" () () ...
  #
  #  ラベル
  #   LABEL 名前（風呂とか洗面とか)
  #
  #  地盤 (主に用壁で用いる)
  #   LEVEL 地盤上げ下げ
  #
  def command2class 
    {
      'I'  => InvertMasu,		# 合流マス
      'ID'  => DansaMasu,	  # 段差付合流マス
      'L'  => LMasu,		    # L桝
      'T'  => TrapMasu,		  # トラップ
      'DT' =>  DoubleTrapMasu,	# ダブルトラップ
      'KT' =>  KitenTrapMasu,	# 起点トラップマス(出の角度を設定できる)
      'D'  => DropMasu,		  # ドロップマス
      'DD' => DoubleDropMasu,	# 2方向ドロップマス
      'LABEL' => Label,		  # ラベル
      'WC' => LabelBennjo,	# 1Fトイレ
      'SM' => LabelSenmen,	# 洗面
      'ST' => LabelSentaku,	# 洗濯
      'SS' => LabelSenmenSentaku, # 洗面・洗濯
      'KI' => LabelDaidokoro,	# 台所
      'BT' => LabelBath,		# 風呂
      'TE' => LabelTearai,	# 手洗
      '2FWC' => Label2FBenjo,	# 2F便所
      '2FSM' => Label2FSenmen,	# 2F洗面

      'TT'    => TTugite,		# チーズ継手 マス無し
      'LT'    => LTugite,		# L継手 マス無し
      'CLT'   => CLTugite,		# 鋳鉄製L継手 マス無し
      'YT'    => YTugite,		# Y継手 マス無し
      'CYT'   => CYTugite,		# 鋳鉄製Y継手 マス無し
      'MC'		=> MCTugite,	# MC継手
      'LA'		=> LATugite,	# ハイパワー
      'VC'		=> VCTugite,	# VCジョイント
      'KC'		=> KCTugite,	# KCジョイント
      'GB'		=> GoodBoxMasu,	# グッドボックス
      'U'		=> GoodBoxMasu,	# グッドボックス(雨水桝)

      'JO'    => Joukasou,		# 浄化槽
      'SO'    => Soshuuki,		# 阻集器(グリストラップ・ランドリートラップなど)
      'PO'    => Pump,				# ポンプ

      'PS'		=> PipeSelect,	# パイプ選択
      'TL'		=> ToridashiLevel,	# 取出しのレベル設定
      'LEVEL' => Level,		  # 地盤高設定(相対指定)
      'ALEVEL' => AbsoluteLevel,    # 地盤高設定(絶対指定 公共マスを0とする)

      'DummyLMasu' => DummyLMasu,
      nil => DummyMasu,		  # ダミー
    }
  end

  def parse_1direction(tree, angle)
    if tree == nil then
      return
    end

    length = tree.car
    unless length.is_a?(Numeric) then
      raise "ここには管の長さを書いて下さい #{tree} #{length}"
    end
    rest = tree.cdr
    masu = nil

    if rest == nil then
      masu = DummyMasu.new
    else
      unless rest.is_a?(List) then
        raise "ここには(桝 長さ ...)のリストを書いて下さい #{rest}"
      end

      masuc = command2class[rest.car]
      if masuc then
        masu = masuc.new
      else
        raise "Unkonw Masu #{rest.car}"
      end

      # マスの名前を設定する
      rest2 = rest.cdr
      rest2 = masu.parse_masuno(rest2)
      # carはマスの名前で多分使わないけど一応残してある
      rest = List.new(rest.car, rest2)
    end

    add_exit(masu, length, angle)

    masu.parse(rest)
  end

  def parse_masuno(tree)
    tree
  end

  def calc_level
    if @enter then
      @enter.calc_level_after
      @level = @enter.level - @enter_length * @koubai
    else
      @level = 0
    end
  end

  def calc_level_after
  end

  def before_recalc_level(sender, prevtl, len)
    false
  end

  def recalc_level(sender, prevtl, len)
    if @enter.before_recalc_level(sender, prevtl, len) == false then
      @enter.recalc_level(self, prevtl, len + @enter_length)
      sender.koubai = @koubai
      sender.calc_level
    end
  end

  def before_draw(device)
  end

  def after_draw(device)
  end

  def addvect
  end

  Color_Table = {
    40 => "2",
    50 => "3",
    65 => "4",
    75 => "5",
    100 => "6",
  }
  def draw(device)
    before_draw(device)

    # パイプを書く
    if @enter then
      entpos = @enter.pos
      lin = draw_line(device, @pos.scalex,
                      @pos.scaley,
                      entpos.scalex,
                      entpos.scaley)
      #      p enter_pipe_size
      #      lin.Cells("LineColor").Formula = Color_Table[enter_pipe_size]
    end

    @exit.each do |n|
      n[1].draw(device)
    end

    after_draw(device)
  end

  def sekisan_parts
    nil
  end

  #	def sekisan_pipe
  #		nil
  #	end
  #  もし、枝管も集計したいならここのコメントを外して
  # sekisan_pipeをコメントアウトする 2004/8/11

  #  LevelやTLを考えると全部集計した方が良い
  #  枝管だった場合は別の方法で除去する
  include SekisanPipeMainParts

  def sekisan
    if enter_pipe_new == "NEW" then

      mname = sekisan_parts
      if mname then
        used_parts[mname] += 1
        $pipe150 += (@level - 0.1)
      end

      sekisan_pipe
    end

    @exit.each do |n|
      n[1].sekisan
    end
  end
end

class MasuCommon<Parts
  @@no = 1

  def assign_no
    if enter_pipe_new == "NEW" and not defined?(@no) then
      @no = @@no
      @@no += 1
    end
  end

  def set_no(n)
    @no = n
  end

  def parse_masuno(tree)
    if tree.car.kind_of?(String) then
      @no = tree.car
      tree.cdr
    else
      tree
    end
  end
end

class KoukyouMasu<MasuCommon

  def initialize
    super

    @angle = 0
    @pos = GeoPoint.new(0, 0)
    @enter = nil
    @pipe_size = 100
    @pipe_kind = $default_material
    #		@pipe_kind = 'VP'
    @pipe_new  = 'NEW'
  end

  def name
    "公共マス"
  end

  include SekisanPipeMainParts

  def parse(tree)
    info = tree[1]
    forward = tree[2]
    right = tree[3]
    left = tree[4]

    if $scale then
      @@scale = $scale
    else
      @@scale = info[1]
    end
    @level = info[0]
    if @level == 0 then
      $hdrawf = false
      $savedir = $conf_savedir
    end
    eangle = 0
    if info.cdr.cdr then
      @angle = info[2]

      if info.cdr.cdr.cdr then
        eangle = info[3]
      end
    end

    parse_1direction(forward, eangle)
    parse_1direction(right, (eangle + 270) % 360)
    parse_1direction(left, (eangle + 90) % 360)

    assign_no
    self
  end

  def before_draw(device)
    device.set_scale(@@scale)
  end

  def next_koubai(k)
    case k
    when 0.02
      0.015
    when 0.015
      0.01
    when 0.01
      0.0075
    when 0.0075
      0.005
    when 0.005
      0.0025
    else
      k
    end
  end

  def recalc_level(sender, prevtl, len)
    if prevtl then
      sender.koubai += (prevtl.level - prevtl.tori_level) / len
      sender.calc_level
      #			print "Recalc koubai in koukyo #{@koubai} -> #{sender.koubai}\n"
      @koubai = sender.koubai
    else
      sender.koubai = next_koubai(sender.koubai)
      sender.calc_level
      print "Recalc executed! #{@koubai} -> #{sender.koubai}\n"
      @koubai = sender.koubai
    end
  end

  def after_draw(device)
    x = @pos.scalex
    y = @pos.scaley
    rad = ((@angle + 180) * Math::PI) / 180.0
    x1 = x + 20 * Math.cos(rad)
    y1 = y + 20 * Math.sin(rad)

    sh = draw_line(device, x, y, x1, y1)
    sh.Cells("EndArrowSize").Formula = "2"
    sh.Cells("EndArrow").Formula = "1"


    draw_oval(device, x - 6, y - 6, x + 6, y + 6)
    draw_oval(device, x - 4, y - 4, x + 4, y + 4)

    if $hdrawf then
      text = "No.%s H%2.2f GL %1.1f %s" % [@no, @level, @alevel, name]
    else
      text = "No.%s GL %1.1f %s" % [@no, @alevel, name]
    end
    draw_label(device, text)

    sh = device.draw_text(0, 0, "1/#{@@scale}", 0, 12)
    #		device.set_font(sh, 12)
  end
end

class DropKoukyouMasu<KoukyouMasu
  def after_draw(device)
    super
    drop_draw(device, 9)
  end
end

class NoneKoukyouMasu<KoukyouMasu
  def after_draw(device)
  end
end

#  2004/10/5 分岐を90°以外の指定が出来るようにした
#  I 角度 右 左
#   角度が数字のときはノーマルを0としたオフセットとする
#
module TreeWayParser
  def eda_size(tree)
    if tree then
      if tree[0] == 0 and tree[1] == "PS" then
        return tree[2].car
      end
    end

    return nil
  end

  def parse_aux(tree)
    angoff = 0
    right = tree[1]
    left = tree[2]
    forward = tree.cdr.cdr.cdr

    # 第1引数が数字ときはoffsetになる
    if right.is_a?(Numeric) then
      angoff = right
      right = tree[2]
      left = tree[3]
      forward = tree.cdr.cdr.cdr.cdr
    end

    unless right.is_a?(List) or right == nil then
      raise "ここには右分岐のリストを書いて下さい #{tree} "
    end

    unless left.is_a?(List) or left == nil then
      raise "ここには左分岐のリストを書いて下さい #{tree} "
    end

    if forward == nil then
      p tree
    end

    langle = 0
    if forward and forward.car.is_a?(List) then
      langle = forward.car.car
      forward = forward.cdr
    end

    @connected_num = 1
    @connected_num += 1 if forward
    @connected_num += 1 if left
    @connected_num += 1 if right

    parse_1direction(forward, langle)
    parse_1direction(left, 90 + angoff)
    lsize = eda_size(left)
    if lsize then
      eda_size = lsize
    end
    parse_1direction(right, 270 + angoff)
    rsize = eda_size(right)
    if rsize then
      eda_size = rsize
    end

    eda_size
  end
end

class TreeWayMasu<MasuCommon
  include TreeWayParser

  def parse(tree)
    parse_aux(tree)

    assign_no
    self
  end

  def connected_num
    @connected_num
  end

  include SekisanPipeMainParts
  def sekisan_parts
    lstep = ((@level + @alevel) * 5 + 0.9999).to_i / 5.0
    ["#{name}, φ150×#{@pipe_size}×H#{lstep}", lstep]
  end
end

class InvertMasu<TreeWayMasu

  def name
    #		"インバートマス"
    #   "Tマス"
    "合流マス"
  end

  def calc_level
    @level = @enter.level - @enter_length * @koubai
  end

  def after_draw(device)
    vanilla_masu_draw(device)
  end

end

class DansaMasu<InvertMasu

  def name
    "合流マス(段差付)"
  end

  def calc_level
    @level = @enter.level - @enter_length * @koubai - 0.05
  end
end

class GoodBoxMasu<TreeWayMasu

  def name
    "雨水桝"
  end

  def calc_level
    @level = @enter.level - @enter_length * @koubai - 0.05
  end

  def after_draw(device)
    sh = box_draw(device, 7, 7)
    draw_label(device, "雨水桝")
  end

end

class TrapMasu<TreeWayMasu

  def name
    "トラップマス"
  end

  def after_draw(device)
    vanilla_masu_draw(device)
    sh = draw_oval(device, @pos.scalex - 1,
                   @pos.scaley - 1,
                   @pos.scalex + 1,
                   @pos.scaley + 1)
    sh.Cells("FillPattern").Formula = "1"
    sh.Cells("FillBkgnd").Formula = "0"
    sh.Cells("FillForegnd").Formula = "0"
  end
end

class DoubleTrapMasu<TrapMasu
  def name
    "ダブルトラップマス"
  end

  def after_draw(device)
    vanilla_masu_draw(device)

    rad = ((@angle + 180) * Math::PI) / 180.0

    [1, -1].each do |d|
      dx = d * Math::cos(rad) * 2
      dy = d * Math::sin(rad) * 2
      sh = draw_oval(device, @pos.scalex + dx + 0.5,
                     @pos.scaley + dy + 0.5,
                     @pos.scalex + dx - 0.5,
                     @pos.scaley + dy - 0.5)
      sh.Cells("FillPattern").Formula = "1"
      sh.Cells("FillBkgnd").Formula = "0"
      sh.Cells("FillForegnd").Formula = "0"
    end
  end
end

class KitenTrapMasu<TrapMasu

  def parse(tree)
    @langle = tree[1]
    forward = tree.cdr.cdr

    parse_1direction(forward, @langle)

    assign_no

    self
  end
end

class Joukasou<Parts
  def name
    "浄化槽"
  end

  include SekisanPipeMainParts
  def sekisan_parts
    lstep = ((@level + @alevel) * 5 + 0.9999).to_i / 5.0
    ["浄化槽", lstep]
  end

  def parse(tree)
    forward = tree.cdr

    parse_1direction(forward, 0)

    self
  end

  def after_draw(device)
    box_draw(device, 15, 30)
    draw_label(device, "浄化槽")

    sh
  end

  def set_position(enter, length, angle)
    off = GeoPoint.new(length + 2, 0).roll(angle, GeoPoint::ORIGIN)

    @pos = enter.pos + off
    @enter = enter
    @enter_length = length
    @angle = angle
  end
end

class Soshuuki<Parts
  def name
    @label
  end

  include SekisanPipeMainParts
  def sekisan_parts
    lstep = ((@level + @alevel) * 5 + 0.9999).to_i / 5.0
    [@label, lstep]
  end

  def parse(tree)
    @label = tree[1]
    forward = tree.cdr.cdr

    parse_1direction(forward, 0)

    self
  end

  XS = 15
  YS = 20

  def after_draw(device)
    x = @pos.scalex
    y = @pos.scaley
    sh = draw_rectangle(device, x, y, XS, YS)
    sh.Cells("Angle").Formula = ((@angle + 90.0) % 360).to_s  + " deg"

    xs2 = XS / 2.0
    ys2 = YS / 2.0
    rad = ((@angle + 90) * Math::PI) / 180.0

    x1 = x + xs2 * Math.cos(rad)
    y1 = y + xs2 * Math.sin(rad)
    sh = draw_line(device, x, y, x1 , y1)

    x1 = x - xs2 * Math.cos(rad)
    y1 = y - xs2 * Math.sin(rad)
    sh = draw_line(device, x, y, x1 , y1)

    draw_label(device, @label)

    sh
  end
end

class Pump<Parts
  def name
    "ポンプ"
  end

  include SekisanPipeMainParts
  def sekisan_parts
    lstep = ((@level + @alevel) * 5 + 0.9999).to_i / 5.0
    [name, lstep]
  end

  def level
    @level + @diff
  end

  def parse(tree)
    @diff = tree[1].to_f
    if @diff == 0 then
      @diff = 0.5
    end
    forward = tree.cdr.cdr

    parse_1direction(forward, 0)

    self
  end

  def after_draw(device)
    x = @pos.scalex
    y = @pos.scaley
    draw_oval(device, x - 7, y - 7, x + 7, y + 7)
    device.draw_text(x, y, "P", @angle)

    draw_label(device, name)
    draw_pipelen(device)
  end
end

class LMasu<MasuCommon

  include SekisanPipeMainParts
  def sekisan_parts
    lstep = ((@level + @alevel) * 5 + 0.9999).to_i / 5.0
    ["#{name} , φ150×#{@pipe_size}×H#{lstep}", lstep]
  end

  def name
    if @langle == 0 then
      "ストレートマス"
    else
      pan = @langle
      if pan > 180 then
        pan = 360 - pan
      end
      if pan < 0 then
        pan = -pan
      end
      "#{pan}° L"
    end
  end

  def parse(tree)
    @langle = tree[1]
    forward = tree.cdr.cdr

    parse_1direction(forward, @langle)

    assign_no

    self
  end

  def after_draw(device)
    vanilla_masu_draw(device)
  end
end

class DropMasu<MasuCommon
  include SekisanPipeMainParts
  def sekisan_parts
    lstep = ((@level + @alevel) * 5 + 0.9999).to_i / 5.0
    ["ドロップマス, φ150×#{@pipe_size}×H#{lstep}", lstep]
  end

  def name
    "ドロップマス"
  end

  def parse(tree)
    parm = tree[1]
    @step = nil
    if parm.kind_of?(List) then
      @step = parm[1]
      @langle = parm[0]
    else
      @langle = parm
    end
    forward = tree.cdr.cdr

    parse_1direction(forward, @langle)

    assign_no

    self
  end

  def level
    if @step then
      @level - @step
    else
      (@level * 1) / 2
    end
  end

  def after_draw(device)
    drop_draw(device, 7)
    vanilla_masu_draw(device)
  end
end

class DoubleDropMasu<DropMasu
  include SekisanPipeMainParts
  def sekisan_parts
    lstep = ((@level + @alevel) * 5 + 0.9999).to_i / 5.0
    ["ダブルドロップマス, φ150×#{@pipe_size}×H#{lstep}", lstep]
  end

  def parse(tree)
    @langle = tree[1]
    right = tree[2]
    left = tree[3]
    forward = tree.cdr.cdr.cdr.cdr

    parse_1direction(forward, @langle)
    parse_1direction(right, @langle - 90)
    parse_1direction(left, @langle + 90)

    assign_no

    self
  end
end

class TugiteCommon<Parts
  def initialize
    @eda_size = nil

    super
  end

  def sekisan_parts
    p @eda_size
    if @eda_size and enter_pipe_size != @eda_size then
      ["#{name}, #{enter_pipe_kind} φ#{enter_pipe_size}×#{@eda_size}", 0]
    else
      ["#{name}, #{enter_pipe_kind} φ#{enter_pipe_size}", 0]
    end
  end

  def parse(tree)
    @label = tree[1]

    parse_1direction(tree.cdr, 0)

    self
  end

  def after_draw(device)
    if $scale or enter_pipe_size >= 100 then
      draw_pipelen(device)
    end

    if $scale then
      draw_label(device, "#{enter_pipe_kind} #{name} φ#{enter_pipe_size}")
    end
  end
end

class LTugite<TugiteCommon

  def name
    pan = @langle
    if pan > 180 then
      pan = 360 - pan
    end
    if pan < 0 then
      pan = -pan
    end

    "#{pan}°L"
  end

  def parse(tree)
    @langle = tree[1]
    forward = tree.cdr.cdr

    parse_1direction(forward, @langle)

    self
  end
end

class CLTugite<LTugite
  def name
    pan = @langle
    if pan > 180 then
      pan = 360 - pan
    end
    if pan < 0 then
      pan = -pan
    end

    "#{pan}°鋳鉄製L"
  end
end

class TTugite<TugiteCommon
  include TreeWayParser

  def parse(tree)
    @eda_size = parse_aux(tree)

    self
  end

  def name
    "T"
  end
end

class YTugite<TugiteCommon
  include TreeWayParser

  def sekisan_parts
    if enter_pipe_size == @eda_size then
      ["#{name}, #{enter_pipe_kind} φ#{enter_pipe_size}", 0]
    else
      ["#{name}, #{enter_pipe_kind} φ#{enter_pipe_size}×#{@eda_size}", 0]
    end
  end


  def name
    pan = @langle
    if pan > 180 then
      pan = 360 - pan
    end
    if pan < 0 then
      pan = -pan
    end

    "#{pan}°YT"
  end

  def parse(tree)
    @connected_num = 1
    @langle = tree[1]
    @eda_size = enter_pipe_size
    forwardr = tree.cdr.cdr
    forwardl = forwardr.cdr
    rest = forwardl.cdr

    @connected_num += 1
    parse_1direction(rest, 0)

    tree = forwardr.car
    if  tree then
      if tree[0] == 0 and tree[1] == "PS" then
        @eda_size = tree[2].car
      end
      cont = List.new(-@langle + 45, tree)
      cont = List.new('DummyLMasu', cont)
      cont = List.new(0.1, cont)
      parse_1direction(cont, -45)
      @connected_num += 1
    end

    tree = forwardl.car
    if tree then
      if tree[0] == 0 and tree[1] == "PS" then
        @eda_size = tree[2].car
      end
      cont = List.new(@langle - 45, tree)
      cont = List.new('DummyLMasu', cont)
      cont = List.new(0.1, cont)
      parse_1direction(cont, 45)
      @connected_num += 1
    end

    self
  end
end

class CYTugite<YTugite
  def name
    pan = @langle
    if pan > 180 then
      pan = 360 - pan
    end
    if pan < 0 then
      pan = -pan
    end

    "#{pan}°鋳鉄製YT"
  end
end

class MCTugite<TugiteCommon
	def name
		"MC"
	end
end

class LATugite<TugiteCommon
  def name
    "LA"
  end
end

class KCTugite<TugiteCommon
  def name
    "KC"
  end
end

class VCTugite<TugiteCommon
  def name
    "VC"
  end
end

class NotPrintParts<Parts
end

class Label<NotPrintParts
  def parse(tree)
    @label = tree[1]

    parse_1direction(tree.cdr.cdr, @angle)

    self
  end

  def addvect
    off = 0.25 + @label.size * 0.0005 * Math.sin(@angle).abs * scale
    pos = @pos.move(off, @angle)
    spos = pos.move(-0.15 * @label.size, 0)
    epos = pos.move(0.15 * @label.size, 0)
    add_vector(spos, epos)
#    p @label.size
#    p @label
=begin
    draw_line(device, spos.scalex, spos.scaley, epos.scalex, epos.scaley)
=end
  end

  def after_draw(device)
    off = 0.25 + @label.size * 0.0005 * Math.sin(@angle).abs * scale
    pos = @pos.move(off, @angle)
    device.draw_text(pos.scalex, pos.scaley, @label, 0)
    # pipelen
    # 12pt(枝の寸法が必要なときコメントアウトをとる)
    # draw_pipelen(device)
  end
end

class FixLabel<Label

  def parse(tree)
    set_label

    parse_1direction(tree.cdr, @angle)

    self
  end
end

class LabelBennjo<FixLabel

  def set_label
    @label = "便所"
  end
end

class LabelSenmen<FixLabel

  def set_label
    @label = "洗面"
  end
end

class Label2FSenmen<FixLabel

  def set_label
    @label = "2F洗面"
  end
end

class LabelSentaku<FixLabel

  def set_label
    @label = "洗濯"
  end
end

class LabelSenmenSentaku<FixLabel

  def set_label
    @label = "洗面・洗濯"
  end
end

class LabelTearai<FixLabel

  def set_label
    @label = "手洗"
  end
end

class Label2FBenjo<FixLabel

  def set_label
    @label = "2F便所"
  end
end

class LabelDaidokoro<FixLabel

  def set_label
    @label = "台所"
  end
end

class LabelBath<FixLabel

  def set_label
    @label = "風呂"
  end
end

class Level<NotPrintParts

  def parse(tree)
#    @step = tree[1]
    @alevel += tree[1].to_f

    parse_1direction(tree.cdr.cdr, 0)

    self
  end

  def after_draw(device)
    if $scale or enter_pipe_size >= 100 then
      draw_pipelen(device)
    end
  end

  def level
    @level
  end

  def calc_level
    @level = @enter.level - @enter_length * @koubai
  end
end

class AbsoluteLevel<NotPrintParts
  def parse(tree)
    @alevel = tree[1].to_f

    parse_1direction(tree.cdr.cdr, 0)

    self
  end

  def after_draw(device)
    #		if $scale or enter_pipe_size >= 100 then
    draw_pipelen(device)
    #		end
  end

  def calc_level
    @level = @enter.level - @enter_length * @koubai
  end
end

class PipeSelect<NotPrintParts

  def parse(tree)
    @old_pipe_size = @pipe_size
    pipe = tree[1]
    @pipe_size = pipe[0]
    p pipe
    if pipe.cdr then
      @pipe_kind = pipe[1]
      if pipe.cdr.cdr then
        @pipe_new = pipe[2]
      end
    end

    parse_1direction(tree.cdr.cdr, 0)

    self
  end

  def sekisan_parts
    if @old_pipe_size != @pipe_size then
      mx = [@old_pipe_size, @pipe_size].max
      mn = [@old_pipe_size, @pipe_size].min
      ["S, φ#{mx}×#{mn}", 0]
    else
      nil
    end
  end

  def after_draw(device)
    if $scale or enter_pipe_size >= 100 then
      draw_pipelen(device)
    end
  end
end

class ToridashiLevel<NotPrintParts

  def parse(tree)
    @tori_level = tree[1]
    parse_1direction(tree.cdr.cdr, 0)

    self
  end

  attr :tori_level

  def calc_level_after
    @level = @enter.level - @enter_length * @koubai
    if (@level - @tori_level).abs > 0.001 then
      @enter.recalc_level(self, self, @enter_length)
    end
    print "Level: #{@level} #{@tori_level} \n"
  end

  def before_recalc_level(sender, prevtl, len)
    # 取出しレベルが設定されている場合は勾配を細かくあわせる
    if prevtl then
      print "TORI: #{prevtl.tori_level} #{tori_level} \n"
      old = sender.koubai
      sender.koubai += (prevtl.level - prevtl.tori_level) / len
      sender.calc_level
    else
      @enter.recalc_level(self, self, len + @enter_length)
      sender.koubai = @koubai
      sender.calc_level
    end

    true
  end

  def after_draw(device)
    if $scale or enter_pipe_size >= 100 then
      draw_pipelen(device)
    end
  end
end

class DummyMasu<NotPrintParts
  def parse(tree)
    self
  end
end

class DummyLMasu<NotPrintParts
  def parse(tree)
    @langle = tree[1]
    forward = tree.cdr.cdr

    parse_1direction(forward, @langle)

    self
  end
end

if __FILE__ == $0 then
  pa = Parser.new

  rt = nil

  $default_material = "VU"
  $hdrawf = true
  $drawf = true
  $sekisanf = true
  $scale = nil
  $sekisansubpipe = false
  $adjust_masulen = false

  argv = ARGV[0].dup # .force_encoding('cp932')
  argv = ARGV[0].dup.encode('cp932')
  while (/\"?^(.*)\.ge~?\"?$/ =~ argv) == nil do
    case argv
    when /--sekisan-sub-pipe/
      # この機能は現在実現されていない。2004/8/11 で検索するといいかも
      $sekisansubpipe = true

    when /--output-dir=(.*)/
      $savedir = $1 + '/'

    when /--no-height-draw/
      $hdrawf = false

    when /--no-draw/
      $drawf = false

    when /--no-sekisan/
      $sekisanf = false

    when /--scale=(\d+)/
      $scale = $1.to_i

    when /--adjust-masu-len/
      $adjust_masulen = true
    end

    ARGV.shift
    argv = ARGV[0].dup.encode('cp932', 'utf-8')
#    argv = ARGV[0].dup
    p argv
  end
  infile = ARGV[0].dup.encode('cp932', 'utf-8').gsub(/\"/, "")
  outfile = File.basename(infile, ".*") + ".vsd"
  #  outfile.force_encoding('sjis')
  #  p outfile.encoding

  p infile
  File.open(infile.encode('utf-8'), "r") do |fp|
    rt = pa.parse(Reader.new(fp).read)
  end

  if $sekisanf then
    rt.sekisan
    wr = ExcelMitumorishoWriter.new
    wr.write(rt.used_parts, rt.used_pipe)
  end

  if $drawf then
    visio = VISIO::VisioDevice.instance
    if rt.init_resolver == nil then
      print "Wanning --- Label conflict happened. I use naive algorithm.\n"
    end
    rt.draw(visio)

    savef = $savedir + outfile
    if test(?e, savef) and File.mtime(infile) < File.mtime(savef) then
      savef = savef + "~"
    end
    p savef
    visio.save_as(savef)
  end

  print $pipe150, "\n"
end  # __FILE__ == $0

