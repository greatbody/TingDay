Attribute VB_Name = "TRC"
Option Explicit

'自定义歌词文件

Type TRCINFOR
    Title As String '歌曲名称
    Singer As String '演唱者
    Album As String '专辑
    From As String '来源
    LrcOffset As String  '时间偏移量 单位 毫秒
    Greetings As String '祝福语（TRC特有 在TingDay播放器上显示）
End Type
