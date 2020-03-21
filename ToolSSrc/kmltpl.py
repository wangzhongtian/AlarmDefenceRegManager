#-*- coding: UTF-8

# format(row["序号"],str2Float(row["经度"]),str2Float(row["纬度"]),
# 					row["牌号"],row["岸别"]+row["GPS测试"] ,row["地名"]  )

g_kmlPosStrfmt_3 ="""
<Placemark>
<name>序{0:s}{4:s}岸  {5:s}  </name>
<description>
<font size="3"></font>
<p>主要信息 </p> <b><font style="color:#00a000">牌号{3:s}#  </font></b>
</description>
<styleUrl>#dot_110_100</styleUrl>
<Point>
<coordinates>{1:f},{2:f}</coordinates>
</Point>
</Placemark>

"""
g_kmlPosStrfmt_一体化 ="""
<Placemark>
<name>序{0:s} </name>
<description>
<font size="3"></font>
<p>间隔距离(围栏片):</p> <b><font style="color:#00a000">{1:f} {5:s}岸{6:s} 牌号{4:s}# 光程{7:s} </font></b>
</description>
<styleUrl>#dot_110_100</styleUrl>
<Point>
<coordinates>{2:f},{3:f}</coordinates>
</Point>
</Placemark>

"""
g_kmlPosStrfmt_2 ="""
<Placemark>
<name>序{0:s} </name>
<description>
<font size="3"></font>
<p>间隔距离(围栏片):</p> <b><font style="color:#00a000">{1:f} {5:s}岸 {6:s} 牌号{4:s}# 光程{7:s} __{8:s} </font></b>
</description>
<styleUrl>#dot_110_100</styleUrl>
<Point>
<coordinates>{2:f},{3:f}</coordinates>
</Point>
</Placemark>

"""
g_kmlPosStrfmt_12 ="""
<Placemark>
<name>序{0:s} </name>
<description>
<font size="3"></font>
<p>间隔距离(围栏片):</p> <b><font style="color:#00a000">{1:f} {5:s}岸 {6:s} 牌号{4:s} </font></b>
</description>
<styleUrl>#dot_110_100</styleUrl>
<Point>
<coordinates>{2:f},{3:f}</coordinates>
</Point>
</Placemark>

"""
g_kmlPosStrfmt_1 ="""
<Placemark>
<name>_牌号_{4:s}#</name>
<description>
<font size="3">序号_{0:s}#</font>
<p>间隔距离(围栏片):</p> <b><font style="color:#00a000">{1:f} </font></b>
</description>
<styleUrl>#dot_110_100</styleUrl>
<Point>
<coordinates>{2:f},{3:f}</coordinates>
</Point>
</Placemark>

"""
g_KMlhead ="""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" 
xmlns:gx="http://www.google.com/kml/ext/2.2" 
xmlns:kml="http://www.opengis.net/kml/2.2" 
xmlns:atom="http://www.w3.org/2005/Atom">
<Document>
<name>Spectrum Logging</name>
<open>1</open>
<StyleMap id="msn_dot_130_bad">
<Pair>
<key>normal</key>
<styleUrl>#dot_130_bad </styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_130_bad _hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_130_bad ">
<IconStyle>
<color>ff0000ff</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_130_bad _hot">
<IconStyle>
<color>ff0000ff</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>


<StyleMap id="msn_dot_130_120">
<Pair>
<key>normal</key>
<styleUrl>#dot_130_120 </styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_130_120 _hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_130_120 ">
<IconStyle>
<color>ff0088ff</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_130_120 _hot">
<IconStyle>
<color>ff0088ff</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>



<StyleMap id="msn_dot_120_110">
<Pair>
<key>normal</key>
<styleUrl>#dot_120_110 </styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_120_110 _hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_120_110 ">
<IconStyle>
<color>ff00ffff</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_120_110 _hot">
<IconStyle>
<color>ff00ffff</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>




<StyleMap id="msn_dot_110_100">
<Pair>
<key>normal</key>
<styleUrl>#dot_110_100 </styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_110_100 _hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_110_100 ">
<IconStyle>
<color>ff00ff88</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_110_100 _hot">
<IconStyle>
<color>ff00ff88</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>



<StyleMap id="msn_dot_100_90">
<Pair>
<key>normal</key>
<styleUrl>#dot_100_90 </styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_100_90 _hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_100_90 ">
<IconStyle>
<color>ffffff00</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_100_90 _hot">
<IconStyle>
<color>ffffff00</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>



<StyleMap id="msn_dot_90_80">
<Pair>
<key>normal</key>
<styleUrl>#dot_90_80 </styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_90_80 _hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_90_80 ">
<IconStyle>
<color>ffff8800</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_90_80 _hot">
<IconStyle>
<color>ffff8800</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>




<StyleMap id="msn_dot_80_fine">
<Pair>
<key>normal</key>
<styleUrl>#dot_80_fine </styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_80_fine _hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_80_fine ">
<IconStyle>
<color>ffff0000</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_80_fine _hot">
<IconStyle>
<color>ffff0000</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<StyleMap id="msn_dot_40_NEAR">
<Pair>
<key>normal</key>
<styleUrl>#dot_40_NEAR</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_40_NEAR _hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_40_NEAR">
<IconStyle>
<color>ffff0000</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_40_NEAR_hot">
<IconStyle>
<color> ffff0000</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>


<StyleMap id="msn_dot_30_40">
<Pair>
<key>normal</key>
<styleUrl>#dot_30_40</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_30_40_hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_130_40">
<IconStyle>
<color> ffff8800</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_30_40_hot">
<IconStyle>
<color> ffff8800</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>



<StyleMap id="msn_dot_20_30">
<Pair>
<key>normal</key>
<styleUrl>#dot_20_30</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_20_30_hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_20_30">
<IconStyle>
<color> ffffff00</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_20_30_hot">
<IconStyle>
<color> ffffff00</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>




<StyleMap id="msn_dot_10_20">
<Pair>
<key>normal</key>
<styleUrl>#dot_10_20</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot_10_20_hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot_10_20">
<IconStyle>
<color> ff00ff44</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot_10_20_hot">
<IconStyle>
<color> ff00ff44</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>




<StyleMap id="msn_dot-10_10">
<Pair>
<key>normal</key>
<styleUrl>#dot-10_10</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot-10_10_hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot-10_10">
<IconStyle>
<color> ff00ff88</color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot-10_10_hot">
<IconStyle>
<color> ff00ff88</color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>





<StyleMap id="msn_dot-20-10">
<Pair>
<key>normal</key>
<styleUrl>#dot-20-10</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot-20-10_hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot-20-10">
<IconStyle>
<color> ff00ffbb </color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot-20-10_hot">
<IconStyle>
<color> ff00ffbb </color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>







<StyleMap id="msn_dot-30-20">
<Pair>
<key>normal</key>
<styleUrl>#dot-30-20</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot-30-20_hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot-30-20">
<IconStyle>
<color> ff00ffff </color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot-30-20_hot">
<IconStyle>
<color> ff00ffff </color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>



<StyleMap id="msn_dot-40-30">
<Pair>
<key>normal</key>
<styleUrl>#dot-40-30</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot-40-30_hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot-40-30">
<IconStyle>
<color> ff0088ff </color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot-40-30_hot">
<IconStyle>
<color> ff0088ff </color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>




<StyleMap id="msn_dot-40_far">
<Pair>
<key>normal</key>
<styleUrl>#dot-40_far</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#dot-40_far_hot</styleUrl>
</Pair>
</StyleMap>

<Style id="dot-40_far">
<IconStyle>
<color> ff0000ff </color>
<scale>0.8</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>

<Style id="dot-40_far_hot">
<IconStyle>
<color> ff0000ff </color>
<scale>1.0</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/paddle/grn-blank-lv.png</href>
</Icon>
</IconStyle>
<ListStyle>
</ListStyle>
</Style>
<Folder>

"""
	
g_KMlTail ="""
</Folder>
</Document>
</kml>
<!-- Ver. 4.5 -->
		"""

g_trackHead="""
<Placemark>
<name>主机轨迹</name>
<styleUrl>#m_ylw-pushpin</styleUrl>
<LineString>
<tessellate>1</tessellate>
<coordinates>
		"""
g_trackTail="""
</coordinates>
</LineString>
</Placemark>
		"""
g_KMlhead_guiji="""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2"
 xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">
<Document>
<name>wew.kml</name>
<Style id="s_ylw-pushpin_hl">
<IconStyle>
<scale>1.3</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>
</Icon>
<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>
</IconStyle>
</Style>
<StyleMap id="m_ylw-pushpin">
<Pair>
<key>normal</key>
<styleUrl>#s_ylw-pushpin</styleUrl>
</Pair>
<Pair>
<key>highlight</key>
<styleUrl>#s_ylw-pushpin_hl</styleUrl>
</Pair>
</StyleMap>
<Style id="s_ylw-pushpin">
<IconStyle>
<scale>1.1</scale>
<Icon>
<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>
</Icon>
<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>
</IconStyle>
</Style>
<Placemark>
<name>未命名路径</name>
<styleUrl>#m_ylw-pushpin</styleUrl>
<LineString>
<tessellate>1</tessellate>
<coordinates>
		"""
g_KMlTailguiji ="""
</coordinates>
</LineString>
</Placemark>
</Document>
</kml>
		"""