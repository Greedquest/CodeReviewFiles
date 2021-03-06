Attribute VB_Name = "myProject"
Option Explicit
Private Type codeItem
    extension As String
    module_name As String
    code_content() As String
End Type

Private Const TypeBinary = 1, vbext_pp_none = 0
Private Const ForReading = 1, ForWriting = 2, ForAppending = 8

Private Function getCodeDefinition(itemNo As Long) As codeItem
    With getCodeDefinition
        Select Case itemNo
            Case 1
                .extension = ".cls"
                .module_name = "StopwatchResults"
                ReDim .code_content(0 To 0)
                .code_content(0) = "VkVSU0lPTiAxLjAgQ0xBU1MNCkJFR0lODQogIE11bHRpVXNlID0gLTEgICdUcnVlDQpFTkQNCkF0dHJpYnV0ZSBWQl9OYW1lID0gIlN0b3B3YXRjaFJlc3VsdHMiDQpBdHRyaWJ1dGUgVkJfR2xvYmFsTmFtZVNwYWNlID0gRmFsc2UNCkF0dHJpYnV0ZSBWQl9DcmVhdGFibGUgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX1ByZWRlY2xhcmVkSWQgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX0V4cG9zZWQgPSBGYWxzZQ0KT3B0aW9uIEV4cGxpY2l0DQoNClByaXZhdGUgVHlwZSBUU3RvcFdhdGNoUmVzdWx0cw0KICAgIFRpbWVEYXRhIEFzIE9iamVjdA0KICAgIExhYmVsRGF0YSBBcyBMYWJlbFRyZWUNCkVuZCBUeXBlDQoNClByaXZhdGUgdGhpcyBBcyBUU3RvcFdhdGNoUmVzdWx0cw0KDQpQdWJsaWMgU3ViIExvYWREYXRhKEJ5VmFsIFRpbWVEYXRhIEFzIE9iamVjdCwgQnlWYWwgTGFiZWxEYXRhIEFzIExhYmVsVHJlZSkNCiAgICBTZXQgdGhpcy5MYWJlbERhdGEgPSBMYWJlbERhdGENCiAgICBTZXQgdGhpcy5UaW1lRGF0YSA9IFRpbWVEYXRhDQogICAgd3JpdGVUaW1lcyB0aGlzLkxhYmVsRGF0YQ0KRW5kIFN1Yg0KDQpQdWJsaWMgUHJvcGVydHkgR2V0IFRvTGFiZWxUcmVlKCkgQXMgTGFiZWxUcmVlDQogICAgU2V0IFRvTGFiZWxUcmVlID0gdGhpcy5MYWJlbERhdGENCkVuZCBQcm9wZXJ0eQ0KDQpQ" & _
"dWJsaWMgUHJvcGVydHkgR2V0IFJhd0RhdGEoKSBBcyBPYmplY3QNCiAgICBTZXQgUmF3RGF0YSA9IHRoaXMuVGltZURhdGENCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgU3ViIFRvSW1tZWRpYXRlV2luZG93KCkNCidQcmludHMgdGltZSBpbmZvIHRvIGltbWVkaWF0ZSB3aW5kb3cNCiAgICBEaW0gcmVzdWx0c1RyZWUgQXMgTGFiZWxUcmVlDQogICAgU2V0IHJlc3VsdHNUcmVlID0gdGhpcy5MYWJlbERhdGENCiAgICBEaW0gZGljdCBBcyBPYmplY3QNCiAgICBTZXQgZGljdCA9IENyZWF0ZU9iamVjdCgiU2NyaXB0aW5nLkRpY3Rpb25hcnkiKQ0KICAgIGZsYXR0ZW5UcmVlIHJlc3VsdHNUcmVlLCBkaWN0DQogICAgRGVidWcuUHJpbnQgIkxhYmVsIG5hbWUiLCAiVGltZSB0YWtlbiINCiAgICBEZWJ1Zy5QcmludCBTdHJpbmcoMzUsICItIikNCiAgICBEaW0gdmFsdWUgQXMgVmFyaWFudA0KICAgIEZvciBFYWNoIHZhbHVlIEluIGRpY3QuS2V5cw0KICAgICAgICBEZWJ1Zy5QcmludCB2YWx1ZSwgZGljdCh2YWx1ZSkoMCksIGRpY3QodmFsdWUpKDEpDQogICAgTmV4dCB2YWx1ZQ0KRW5kIFN1Yg0KDQpQcml2YXRlIFN1YiBmbGF0dGVuVHJlZShCeVZhbCB0cmVlSXRlbSBBcyBMYWJlbFRyZWUsIEJ5UmVmIGRpY3QgQXMgT2JqZWN0LCBPcHRpb25hbCBCeVZhbCBkZXB0aCBBcyBMb25nID0gMCkN" & _
"CidyZWN1cnNpdmVseSBjb252ZXJ0cyBhIHJlc3VsdHMgdHJlZSB0byBhIGRpY3Rpb25hcnkgb2YgcmVzdWx0IGtleXMNCiAgICBkaWN0LkFkZCBwcmludGYoInswfSB7MX0iLCB0cmVlSXRlbS5Mb2NhdGlvbiwgdHJlZUl0ZW0uTm9kZU5hbWUpLCBBcnJheSh0cmVlSXRlbS5UaW1lU3BlbnQsIHRyZWVJdGVtLlRpbWVXYXN0ZWQpDQogICAgSWYgdHJlZUl0ZW0uQ2hpbGROb2Rlcy5Db3VudCA+IDAgVGhlbg0KICAgICAgICBEaW0gaXRlbSBBcyBWYXJpYW50DQogICAgICAgIEZvciBFYWNoIGl0ZW0gSW4gdHJlZUl0ZW0uQ2hpbGROb2Rlcw0KICAgICAgICAgICAgZmxhdHRlblRyZWUgaXRlbSwgZGljdCwgZGVwdGggKyAxDQogICAgICAgIE5leHQNCiAgICBFbmQgSWYNCkVuZCBTdWINCg0KUHJpdmF0ZSBTdWIgd3JpdGVUaW1lcyhCeVJlZiBsYWJlbEl0ZW0gQXMgTGFiZWxUcmVlKQ0KICAgICdSZWN1cnNpdmVseSB3cml0ZSBhYnNvbHV0ZSB0aW1lIGRhdGEgdG8gdGltZSBlbGFwc2VkIGRhdGENCg0KICAgIERpbSBzdGFydFRpbWVzIEFzIFRpbWVJbmZvDQogICAgRGltIGVuZFRpbWVzIEFzIFRpbWVJbmZvDQoNCiAgICBzZXRUaW1lU3RhbXBzIGxhYmVsSXRlbSwgc3RhcnRUaW1lcywgZW5kVGltZXMgJ2dldCB0aW1lc3RhbXBzIGZyb20gZGljdGlvbmFyeQ0KICAgIFdpdGggbGFiZWxJdGVt" & _
"DQogICAgICAgIElmIC5DaGlsZE5vZGVzLkNvdW50ID4gMCBUaGVuDQogICAgICAgICAgICAnaGFzIGNoaWxkcmVuLCB3b3JrIG91dCB0aW1lIHNwZW50IGZvciBlYWNoIHRoZW4gc3VtDQogICAgICAgICAgICBEaW0gY2hpbGRMYWJlbCBBcyBMYWJlbFRyZWUNCiAgICAgICAgICAgIERpbSBpdGVtIEFzIFZhcmlhbnQNCg0KICAgICAgICAgICAgRm9yIEVhY2ggaXRlbSBJbiAuQ2hpbGROb2RlcyAgICAgICAgICdyZWN1cnNlIGRlZXBlcg0KICAgICAgICAgICAgICAgIFNldCBjaGlsZExhYmVsID0gaXRlbQ0KICAgICAgICAgICAgICAgIHdyaXRlVGltZXMgY2hpbGRMYWJlbA0KICAgICAgICAgICAgICAgIC5UaW1lV2FzdGVkID0gLlRpbWVXYXN0ZWQgKyBjaGlsZExhYmVsLlRpbWVXYXN0ZWQgJ2FkZCB1cCBjaGlsZCB3YXN0ZWQgdGltZQ0KICAgICAgICAgICAgTmV4dCBpdGVtDQogICAgICAgICAgICAuVGltZVNwZW50ID0gZW5kVGltZXMuVGltZUluIC0gc3RhcnRUaW1lcy5UaW1lT3V0IC0gLlRpbWVXYXN0ZWQgJ3RpbWUgZGlmZiAtIHdhc3RlZCB0aW1lDQogICAgICAgICAgICAuVGltZVdhc3RlZCA9IC5UaW1lV2FzdGVkICsgZW5kVGltZXMuVGltZU91dCAtIGVuZFRpbWVzLlRpbWVJbiArIHN0YXJ0VGltZXMuVGltZU91dCAtIHN0YXJ0VGltZXMuVGltZUluDQogICAgICAgIEVsc2Ug" & _
"ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgJ05vIGNoaWxkcmVuDQogICAgICAgICAgICBJZiAuTGFiZWxUeXBlID0gc3RwX0xhcFRpbWUgVGhlbg0KICAgICAgICAgICAgICAgIC5UaW1lV2FzdGVkID0gZW5kVGltZXMuVGltZU91dCAtIGVuZFRpbWVzLlRpbWVJbg0KICAgICAgICAgICAgRWxzZSAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICdmaW5kIHRpbWUgc3RhbXBzIGZvciBvcGVuaW5nIGFuZCBjbG9zaW5nIGxhYmVsDQogICAgICAgICAgICAgICAgLlRpbWVXYXN0ZWQgPSBlbmRUaW1lcy5UaW1lT3V0IC0gZW5kVGltZXMuVGltZUluICsgc3RhcnRUaW1lcy5UaW1lT3V0IC0gc3RhcnRUaW1lcy5UaW1lSW4NCiAgICAgICAgICAgIEVuZCBJZg0KICAgICAgICAgICAgLlRpbWVTcGVudCA9IGVuZFRpbWVzLlRpbWVJbiAtIHN0YXJ0VGltZXMuVGltZU91dA0KICAgICAgICBFbmQgSWYNCiAgICBFbmQgV2l0aA0KRW5kIFN1Yg0KDQpQcml2YXRlIFN1YiBzZXRUaW1lU3RhbXBzKEJ5VmFsIGxhYmVsSXRlbSBBcyBMYWJlbFRyZWUsIEJ5UmVmIHN0YXJ0VGltZXMgQXMgVGltZUluZm8sIEJ5UmVmIGVuZFRpbWVzIEFzIFRpbWVJbmZvKQ0KICAgICd3cml0ZXMgdGltZXN0YW1wcyBieXJlZg0KICAgIFdpdGggbGFiZWxJdGVtDQogICAgICAgIERpbSBzdGFydEtl" & _
"eSBBcyBTdHJpbmcNCiAgICAgICAgRGltIGVuZEtleSBBcyBTdHJpbmcNCiAgICAgICAgJ2xvY2F0aW9uIG9mIHRpbWVzdGFtcHMgaW4gZGljdGlvbmFyeQ0KICAgICAgICBTZWxlY3QgQ2FzZSAuTGFiZWxUeXBlDQogICAgICAgIENhc2Ugc3RwX0xhcFRpbWUNCiAgICAgICAgICAgIERpbSBrZXlCYXNlIEFzIFN0cmluZw0KICAgICAgICAgICAga2V5QmFzZSA9IC5wYXJlbnROb2RlLkxvY2F0aW9uDQogICAgICAgICAgICBEaW0gbGFwTnVtYmVyIEFzIExvbmcNCiAgICAgICAgICAgIGxhcE51bWJlciA9IFJpZ2h0JCguTm9kZU5hbWUsIExlbiguTm9kZU5hbWUpIC0gMykNCiAgICAgICAgICAgIElmIGxhcE51bWJlciA9IDEgVGhlbiAgICAgICAgICAgICAgICAnZmlyc3QgbGFwLCBzdGFydHMgYXQNCiAgICAgICAgICAgICAgICBzdGFydEtleSA9IHByaW50ZigiezB9X29wZW4iLCBrZXlCYXNlKQ0KICAgICAgICAgICAgRWxzZQ0KICAgICAgICAgICAgICAgIHN0YXJ0S2V5ID0gcHJpbnRmKCJ7MH1fTGFwezF9Iiwga2V5QmFzZSwgbGFwTnVtYmVyIC0gMSkgJ3N0YXJ0IGF0IHByZXYgbGFwLCBlbmQgaGVyZQ0KICAgICAgICAgICAgRW5kIElmDQogICAgICAgICAgICBlbmRLZXkgPSBwcmludGYoInswfV9MYXB7MX0iLCBrZXlCYXNlLCBsYXBOdW1iZXIpDQogICAgICAgIENhc2UgRWxzZQ0K" & _
"ICAgICAgICAgICAgc3RhcnRLZXkgPSBwcmludGYoInswfV9vcGVuIiwgLkxvY2F0aW9uKQ0KICAgICAgICAgICAgZW5kS2V5ID0gcHJpbnRmKCJ7MH1fY2xvc2UiLCAuTG9jYXRpb24pDQogICAgICAgIEVuZCBTZWxlY3QNCiAgICAgICAgU2V0IGVuZFRpbWVzID0gdGhpcy5UaW1lRGF0YShlbmRLZXkpDQogICAgICAgIFNldCBzdGFydFRpbWVzID0gdGhpcy5UaW1lRGF0YShzdGFydEtleSkNCiAgICBFbmQgV2l0aA0KDQpFbmQgU3ViDQoNClByaXZhdGUgRnVuY3Rpb24gcHJpbnRmKEJ5VmFsIG1hc2sgQXMgU3RyaW5nLCBQYXJhbUFycmF5IHRva2VucygpKSBBcyBTdHJpbmcNCidGb3JtYXQgc3RyaW5nIHdpdGggYnkgc3Vic3RpdHV0aW5nIGludG8gbWFzayAtIHN0YWNrb3ZlcmZsb3cuY29tL2EvMTcyMzM4MzQvNjYwOTg5Ng0KICAgIERpbSBpIEFzIExvbmcNCiAgICBGb3IgaSA9IDAgVG8gVUJvdW5kKHRva2VucykNCiAgICAgICAgbWFzayA9IFJlcGxhY2UkKG1hc2ssICJ7IiAmIGkgJiAifSIsIHRva2VucyhpKSkNCiAgICBOZXh0DQogICAgcHJpbnRmID0gbWFzaw0KRW5kIEZ1bmN0aW9uDQo="
            Case 2
                .extension = ".cls"
                .module_name = "TimeInfo"
                ReDim .code_content(0 To 0)
                .code_content(0) = "VkVSU0lPTiAxLjAgQ0xBU1MNCkJFR0lODQogIE11bHRpVXNlID0gLTEgICdUcnVlDQpFTkQNCkF0dHJpYnV0ZSBWQl9OYW1lID0gIlRpbWVJbmZvIg0KQXR0cmlidXRlIFZCX0dsb2JhbE5hbWVTcGFjZSA9IEZhbHNlDQpBdHRyaWJ1dGUgVkJfQ3JlYXRhYmxlID0gRmFsc2UNCkF0dHJpYnV0ZSBWQl9QcmVkZWNsYXJlZElkID0gRmFsc2UNCkF0dHJpYnV0ZSBWQl9FeHBvc2VkID0gRmFsc2UNCk9wdGlvbiBFeHBsaWNpdA0KDQpQcml2YXRlIFR5cGUgVFRpbWVJbmZvDQogICAgVGltZUluIEFzIERvdWJsZQ0KICAgIFRpbWVPdXQgQXMgRG91YmxlDQpFbmQgVHlwZQ0KDQpQcml2YXRlIHRoaXMgQXMgVFRpbWVJbmZvDQoNClB1YmxpYyBQcm9wZXJ0eSBHZXQgVGltZUluKCkgQXMgRG91YmxlDQogICAgVGltZUluID0gdGhpcy5UaW1lSW4NCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgUHJvcGVydHkgTGV0IFRpbWVJbihCeVZhbCB2YWx1ZSBBcyBEb3VibGUpDQogICAgdGhpcy5UaW1lSW4gPSB2YWx1ZQ0KRW5kIFByb3BlcnR5DQoNClB1YmxpYyBQcm9wZXJ0eSBHZXQgVGltZU91dCgpIEFzIERvdWJsZQ0KICAgIFRpbWVPdXQgPSB0aGlzLlRpbWVPdXQNCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgUHJvcGVydHkgTGV0IFRpbWVPdXQoQnlWYWwgdmFsdWUgQXMgRG91YmxlKQ0KICAgIHRoaXMuVGlt" & _
"ZU91dCA9IHZhbHVlDQpFbmQgUHJvcGVydHkNCg0K"
            Case 3
                .extension = ".cls"
                .module_name = "LabelTree"
                ReDim .code_content(0 To 0)
                .code_content(0) = "VkVSU0lPTiAxLjAgQ0xBU1MNCkJFR0lODQogIE11bHRpVXNlID0gLTEgICdUcnVlDQpFTkQNCkF0dHJpYnV0ZSBWQl9OYW1lID0gIkxhYmVsVHJlZSINCkF0dHJpYnV0ZSBWQl9HbG9iYWxOYW1lU3BhY2UgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX0NyZWF0YWJsZSA9IEZhbHNlDQpBdHRyaWJ1dGUgVkJfUHJlZGVjbGFyZWRJZCA9IEZhbHNlDQpBdHRyaWJ1dGUgVkJfRXhwb3NlZCA9IEZhbHNlDQpPcHRpb24gRXhwbGljaXQNCg0KUHVibGljIEVudW0gc3RvcHdhdGNoTGFibGVUeXBlDQogICAgc3RwX0xhcFRpbWUgPSAxDQogICAgc3RwX0xhYmVsDQogICAgc3RwX1N0YXJ0DQogICAgc3RwX0ZpbmlzaA0KRW5kIEVudW0NCg0KUHJpdmF0ZSBUeXBlIFRMYWJlbFRyZWUNCiAgICBwYXJlbnROb2RlIEFzIExhYmVsVHJlZQ0KICAgIENoaWxkTm9kZXMgQXMgQ29sbGVjdGlvbg0KICAgIE5vZGVOYW1lIEFzIFN0cmluZw0KICAgIFRpbWVTcGVudCBBcyBEb3VibGUNCiAgICBUaW1lV2FzdGVkIEFzIERvdWJsZSAgICAgICAgICAgICAgICAgICAgICAgICAndGltZSB1c2VkIGJ5IHN0b3B3YXRjaCBydW5zDQogICAgTG9jYXRpb24gQXMgU3RyaW5nDQogICAgTGFiZWxUeXBlIEFzIHN0b3B3YXRjaExhYmxlVHlwZQ0KRW5kIFR5cGUNCg0KUHJpdmF0ZSB0aGlzIEFzIFRMYWJlbFRyZWUNClB1YmxpYyBQ" & _
"cm9wZXJ0eSBHZXQgTGFiZWxUeXBlKCkgQXMgc3RvcHdhdGNoTGFibGVUeXBlDQogICAgTGFiZWxUeXBlID0gdGhpcy5MYWJlbFR5cGUNCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgUHJvcGVydHkgTGV0IExhYmVsVHlwZShCeVZhbCB2YWx1ZSBBcyBzdG9wd2F0Y2hMYWJsZVR5cGUpDQogICAgdGhpcy5MYWJlbFR5cGUgPSB2YWx1ZQ0KRW5kIFByb3BlcnR5DQoNClB1YmxpYyBQcm9wZXJ0eSBHZXQgTG9jYXRpb24oKSBBcyBTdHJpbmcNCiAgICBMb2NhdGlvbiA9IHRoaXMuTG9jYXRpb24NCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgUHJvcGVydHkgTGV0IExvY2F0aW9uKEJ5VmFsIHZhbHVlIEFzIFN0cmluZykNCiAgICB0aGlzLkxvY2F0aW9uID0gdmFsdWUNCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgUHJvcGVydHkgR2V0IFRpbWVTcGVudCgpIEFzIERvdWJsZQ0KICAgIFRpbWVTcGVudCA9IHRoaXMuVGltZVNwZW50DQpFbmQgUHJvcGVydHkNCg0KUHVibGljIFByb3BlcnR5IExldCBUaW1lU3BlbnQoQnlWYWwgdmFsdWUgQXMgRG91YmxlKQ0KICAgIHRoaXMuVGltZVNwZW50ID0gdmFsdWUNCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgUHJvcGVydHkgR2V0IFRpbWVXYXN0ZWQoKSBBcyBEb3VibGUNCiAgICBUaW1lV2FzdGVkID0gdGhpcy5UaW1lV2FzdGVkDQpFbmQgUHJvcGVydHkNCg0KUHVi" & _
"bGljIFByb3BlcnR5IExldCBUaW1lV2FzdGVkKEJ5VmFsIHZhbHVlIEFzIERvdWJsZSkNCiAgICB0aGlzLlRpbWVXYXN0ZWQgPSB2YWx1ZQ0KRW5kIFByb3BlcnR5DQoNClB1YmxpYyBQcm9wZXJ0eSBHZXQgQ2hpbGROb2RlcygpIEFzIENvbGxlY3Rpb24NCiAgICBTZXQgQ2hpbGROb2RlcyA9IHRoaXMuQ2hpbGROb2Rlcw0KRW5kIFByb3BlcnR5DQoNClB1YmxpYyBQcm9wZXJ0eSBTZXQgQ2hpbGROb2RlcyhCeVZhbCB2YWx1ZSBBcyBDb2xsZWN0aW9uKQ0KICAgIFNldCB0aGlzLkNoaWxkTm9kZXMgPSB2YWx1ZQ0KRW5kIFByb3BlcnR5DQoNClB1YmxpYyBQcm9wZXJ0eSBHZXQgTm9kZU5hbWUoKSBBcyBTdHJpbmcNCiAgICBOb2RlTmFtZSA9IHRoaXMuTm9kZU5hbWUNCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgUHJvcGVydHkgTGV0IE5vZGVOYW1lKEJ5VmFsIHZhbHVlIEFzIFN0cmluZykNCiAgICB0aGlzLk5vZGVOYW1lID0gdmFsdWUNCkVuZCBQcm9wZXJ0eQ0KDQpQdWJsaWMgUHJvcGVydHkgR2V0IHBhcmVudE5vZGUoKSBBcyBMYWJlbFRyZWUNCiAgICBTZXQgcGFyZW50Tm9kZSA9IHRoaXMucGFyZW50Tm9kZQ0KRW5kIFByb3BlcnR5DQoNClB1YmxpYyBQcm9wZXJ0eSBTZXQgcGFyZW50Tm9kZShCeVZhbCB2YWx1ZSBBcyBMYWJlbFRyZWUpDQogICAgU2V0IHRoaXMucGFyZW50Tm9kZSA9" & _
"IHZhbHVlDQpFbmQgUHJvcGVydHkNCg0KUHJpdmF0ZSBTdWIgQ2xhc3NfSW5pdGlhbGl6ZSgpDQogICAgU2V0IHRoaXMuQ2hpbGROb2RlcyA9IE5ldyBDb2xsZWN0aW9uDQpFbmQgU3ViDQoNCg=="
            Case 4
                .extension = ".cls"
                .module_name = "Stopwatch"
                ReDim .code_content(0 To 0)
                .code_content(0) = "VkVSU0lPTiAxLjAgQ0xBU1MNCkJFR0lODQogIE11bHRpVXNlID0gLTEgICdUcnVlDQpFTkQNCkF0dHJpYnV0ZSBWQl9OYW1lID0gIlN0b3B3YXRjaCINCkF0dHJpYnV0ZSBWQl9HbG9iYWxOYW1lU3BhY2UgPSBGYWxzZQ0KQXR0cmlidXRlIFZCX0NyZWF0YWJsZSA9IEZhbHNlDQpBdHRyaWJ1dGUgVkJfUHJlZGVjbGFyZWRJZCA9IEZhbHNlDQpBdHRyaWJ1dGUgVkJfRXhwb3NlZCA9IEZhbHNlDQpPcHRpb24gRXhwbGljaXQNCg0KUHJpdmF0ZSBUeXBlIFRTdG9wV2F0Y2gNCiAgICBkYXRhIEFzIE9iamVjdA0KICAgIEN1cnJlbnRMYWJlbCBBcyBMYWJlbFRyZWUNCiAgICBSZXN1bHRzIEFzIFN0b3B3YXRjaFJlc3VsdHMNCiAgICBGaXJzdExhYmVsIEFzIExhYmVsVHJlZQ0KRW5kIFR5cGUNCg0KUHJpdmF0ZSB0aGlzIEFzIFRTdG9wV2F0Y2gNCg0KUHJpdmF0ZSBEZWNsYXJlIFB0clNhZmUgRnVuY3Rpb24gZ2V0RnJlcXVlbmN5IExpYiAia2VybmVsMzIiIF8NCkFsaWFzICJRdWVyeVBlcmZvcm1hbmNlRnJlcXVlbmN5IiAoY3lGcmVxdWVuY3kgQXMgQ3VycmVuY3kpIEFzIExvbmcNClByaXZhdGUgRGVjbGFyZSBQdHJTYWZlIEZ1bmN0aW9uIGdldFRpY2tDb3VudCBMaWIgImtlcm5lbDMyIiBfDQpBbGlhcyAiUXVlcnlQZXJmb3JtYW5jZUNvdW50ZXIiIChjeVRpY2tDb3VudCBBcyBDdXJyZW5j" & _
"eSkgQXMgTG9uZw0KDQpQcml2YXRlIEZ1bmN0aW9uIE1pY3JvVGltZXIoKSBBcyBEb3VibGUNCiAgICAnQWNjdXJhdGUgdGltaW5nIG1ldGhvZCAtIHN0YWNrb3ZlcmZsb3cuY29tL2EvNzExNjkyOC82NjA5ODk2DQogICAgRGltIGN5VGlja3MxIEFzIEN1cnJlbmN5DQogICAgU3RhdGljIGN5RnJlcXVlbmN5IEFzIEN1cnJlbmN5DQoNCiAgICBNaWNyb1RpbWVyID0gMA0KDQogICAgSWYgY3lGcmVxdWVuY3kgPSAwIFRoZW4gZ2V0RnJlcXVlbmN5IGN5RnJlcXVlbmN5DQoNCiAgICBnZXRUaWNrQ291bnQgY3lUaWNrczENCg0KICAgIElmIGN5RnJlcXVlbmN5IFRoZW4gTWljcm9UaW1lciA9IGN5VGlja3MxIC8gY3lGcmVxdWVuY3kNCkVuZCBGdW5jdGlvbg0KDQpQdWJsaWMgU3ViIFN0YXJ0KCkNCiAgICBPcGVuTGFiZWwgIlN0YXJ0Ig0KRW5kIFN1Yg0KDQpQdWJsaWMgU3ViIEZpbmlzaCgpDQogICAgQ2xvc2VMYWJlbA0KICAgIFNldCB0aGlzLlJlc3VsdHMgPSBOZXcgU3RvcHdhdGNoUmVzdWx0cw0KICAgIHRoaXMuUmVzdWx0cy5Mb2FkRGF0YSB0aGlzLmRhdGEsIHRoaXMuRmlyc3RMYWJlbA0KRW5kIFN1Yg0KDQpQdWJsaWMgUHJvcGVydHkgR2V0IFJlc3VsdHMoKSBBcyBTdG9wd2F0Y2hSZXN1bHRzDQogICAgU2V0IFJlc3VsdHMgPSB0aGlzLlJlc3VsdHMNCkVuZCBQcm9wZXJ0eQ0KDQpQ" & _
"dWJsaWMgU3ViIE9wZW5MYWJlbChCeVZhbCBsYWJlbE5hbWUgQXMgU3RyaW5nKQ0KICAgICdTYXZlIHRpbWUgb24gYXJyaXZhbA0KICAgIERpbSBjbG9ja1RpbWVzIEFzIE5ldyBUaW1lSW5mbw0KICAgIGNsb2NrVGltZXMuVGltZUluID0gTWljcm9UaW1lcg0KICAgIA0KICAgICdEZWZpbmUgbmV3IGxhYmVsLCBhbmQgbWFrZSBpdCBhIGNoaWxkIG9mIHRoZSBjdXJyZW50IGxhYmVsDQogICAgRGltIG5ld05vZGUgQXMgTmV3IExhYmVsVHJlZQ0KICAgIG5ld05vZGUuTm9kZU5hbWUgPSBsYWJlbE5hbWUNCiAgICBJZiBOb3QgdGhpcy5DdXJyZW50TGFiZWwgSXMgTm90aGluZyBUaGVuDQogICAgICAgIFNldCBuZXdOb2RlLnBhcmVudE5vZGUgPSB0aGlzLkN1cnJlbnRMYWJlbA0KICAgICAgICAnMS4yLjEgZm9ybWF0DQogICAgICAgIG5ld05vZGUuTG9jYXRpb24gPSB0aGlzLkN1cnJlbnRMYWJlbC5Mb2NhdGlvbiAmICIuIiAmIHRoaXMuQ3VycmVudExhYmVsLkNoaWxkTm9kZXMuQ291bnQgKyAxDQogICAgICAgIHRoaXMuQ3VycmVudExhYmVsLkNoaWxkTm9kZXMuQWRkIG5ld05vZGUsIG5ld05vZGUuTG9jYXRpb24gJiBuZXdOb2RlLk5vZGVOYW1lDQogICAgRWxzZQ0KICAgICAgICBuZXdOb2RlLkxvY2F0aW9uID0gIjEiDQogICAgICAgIFNldCB0aGlzLkZpcnN0TGFiZWwgPSBuZXdOb2Rl" & _
"DQogICAgRW5kIElmDQogICAgU2V0IHRoaXMuQ3VycmVudExhYmVsID0gbmV3Tm9kZQ0KICAgIA0KICAgICdTYXZlIHRpbWUgZGF0YSB0byBkaWN0aW9uYXJ5IGFuZCByZXR1cm4gdG8gZXhlY3V0aW9uDQogICAgRGltIGRpY3RLZXkgQXMgU3RyaW5nDQogICAgZGljdEtleSA9IG5ld05vZGUuTG9jYXRpb24gJiAiX29wZW4iDQogICAgdGhpcy5kYXRhLkFkZCBkaWN0S2V5LCBjbG9ja1RpbWVzDQogICAgdGhpcy5kYXRhKGRpY3RLZXkpLlRpbWVPdXQgPSBNaWNyb1RpbWVyDQpFbmQgU3ViDQoNClB1YmxpYyBTdWIgQ2xvc2VMYWJlbCgpDQogICAgJ1NhdmUgdGltZSBvbiBhcnJpdmFsDQogICAgRGltIGNsb2NrVGltZXMgQXMgTmV3IFRpbWVJbmZvDQogICAgY2xvY2tUaW1lcy5UaW1lSW4gPSBNaWNyb1RpbWVyDQogICAgDQogICAgJ1NhdmUgdGltZSBkYXRhIHRvIGRpY3Rpb25hcnkgYW5kIHJldHVybiB0byBleGVjdXRpb24NCiAgICBEaW0gZGljdEtleSBBcyBTdHJpbmcNCiAgICBkaWN0S2V5ID0gdGhpcy5DdXJyZW50TGFiZWwuTG9jYXRpb24gJiAiX2Nsb3NlIg0KICAgIHRoaXMuZGF0YS5BZGQgZGljdEtleSwgY2xvY2tUaW1lcw0KICAgIA0KICAgICdDbG9zZSBsYWJlbCBieSBzZXR0aW5nIHRvIHBhcmVudA0KICAgIFNldCB0aGlzLkN1cnJlbnRMYWJlbCA9IHRoaXMuQ3VycmVudExh" & _
"YmVsLnBhcmVudE5vZGUNCiAgICB0aGlzLmRhdGEoZGljdEtleSkuVGltZU91dCA9IE1pY3JvVGltZXINCkVuZCBTdWINCg0KUHVibGljIFN1YiBMYXAoKQ0KICAgICdTYXZlIHRpbWUgb24gYXJyaXZhbA0KICAgIERpbSBjbG9ja1RpbWVzIEFzIE5ldyBUaW1lSW5mbw0KICAgIGNsb2NrVGltZXMuVGltZUluID0gTWljcm9UaW1lcg0KICAgIA0KICAgICdEZWZpbmUgbmV3IGxhYmVsLCBhbmQgbWFrZSBpdCBhIGNoaWxkIG9mIHRoZSBjdXJyZW50IGxhYmVsDQogICAgRGltIG5ld05vZGUgQXMgTmV3IExhYmVsVHJlZQ0KICAgIG5ld05vZGUuTG9jYXRpb24gPSB0aGlzLkN1cnJlbnRMYWJlbC5Mb2NhdGlvbiAmICIuIiAmIHRoaXMuQ3VycmVudExhYmVsLkNoaWxkTm9kZXMuQ291bnQgKyAxDQogICAgbmV3Tm9kZS5Ob2RlTmFtZSA9ICJMYXAiICYgdGhpcy5DdXJyZW50TGFiZWwuQ2hpbGROb2Rlcy5Db3VudCArIDEgJ3RoaXMuQ3VycmVudExhYmVsLk5vZGVOYW1lICYgIl8NCiAgICBuZXdOb2RlLkxhYmVsVHlwZSA9IHN0cF9MYXBUaW1lDQogICAgDQogICAgSWYgdGhpcy5DdXJyZW50TGFiZWwgSXMgTm90aGluZyBUaGVuDQogICAgICAgIEVyci5EZXNjcmlwdGlvbiA9ICJObyB0ZXN0IGlzIGN1cnJlbnRseSBydW5uaW5nIHRvIHdyaXRlIGxhcCBkYXRhIHRvIg0KICAgICAgICBFcnIuUmFp" & _
"c2UgNQ0KICAgIEVsc2UNCiAgICAgICAgU2V0IG5ld05vZGUucGFyZW50Tm9kZSA9IHRoaXMuQ3VycmVudExhYmVsDQogICAgICAgIHRoaXMuQ3VycmVudExhYmVsLkNoaWxkTm9kZXMuQWRkIG5ld05vZGUsIG5ld05vZGUuTm9kZU5hbWUNCiAgICBFbmQgSWYNCiAgICANCiAgICANCiAgICAnU2F2ZSB0aW1lIGRhdGEgdG8gZGljdGlvbmFyeSBhbmQgcmV0dXJuIHRvIGV4ZWN1dGlvbg0KICAgIERpbSBkaWN0S2V5IEFzIFN0cmluZw0KICAgIGRpY3RLZXkgPSB0aGlzLkN1cnJlbnRMYWJlbC5Mb2NhdGlvbiAmICJfIiAmIG5ld05vZGUuTm9kZU5hbWUNCiAgICB0aGlzLmRhdGEuQWRkIGRpY3RLZXksIGNsb2NrVGltZXMNCiAgICB0aGlzLmRhdGEoZGljdEtleSkuVGltZU91dCA9IE1pY3JvVGltZXINCkVuZCBTdWINCg0KUHJpdmF0ZSBTdWIgQ2xhc3NfSW5pdGlhbGl6ZSgpDQogICAgU2V0IHRoaXMuZGF0YSA9IENyZWF0ZU9iamVjdCgiU2NyaXB0aW5nLkRpY3Rpb25hcnkiKQ0KRW5kIFN1Yg0K"
        Case Else
            .extension = "missing"
        End Select
    End With
End Function

Public Sub Extract()
    Dim code_module As codeItem
    Dim savedPath As String, basePath As String
    Dim i As Long
    'check if vbproject accessible
    If Not project_accessible Then
        MsgBox "The VBA project cannot be accessed programmatically. Ensure programmatic access to Office VBA project is enabled and that the workbook is not protected."
        Exit Sub
    End If
    'check if temp folder acessible
    i = 0
    basePath = Environ$("Temp") & "\"
    Do While True
        i = i + 1
        code_module = getCodeDefinition(i)
        If code_module.extension = "missing" Then
            Exit Do
        Else
            savedPath = createFile(code_module, basePath)
            importFile savedPath
            Kill savedPath
        End If
    Loop
    removemodule "myProject"
End Sub

Private Function project_accessible() As Boolean
    On Error Resume Next
    With ThisWorkbook.VBProject
        project_accessible = .Protection = vbext_pp_none
        project_accessible = project_accessible And Err.Number = 0
    End With
End Function

Private Function createFile(definition As codeItem, filePath As Variant) As String
    Dim codeIndex As Long
    Dim newFileObj As Object
    Set newFileObj = CreateObject("ADODB.Stream")
    newFileObj.Type = TypeBinary
    'Open the stream and write binary data
    newFileObj.Open
    'create file from x64 string
    With definition
        Dim bytes() As Byte
        Dim fullPath As String
        fullPath = filePath & .module_name & .extension
        bytes = FromBase64(Join(.code_content))
        newFileObj.Write bytes
        newFileObj.SaveToFile fullPath, ForWriting
        createFile = fullPath
    End With
End Function

Private Sub importFile(filePath As String)
    ThisWorkbook.VBProject.VBComponents.Import filePath
End Sub

Private Function removemodule(moduleName As String) As Boolean
    On Error Resume Next
    With ThisWorkbook.VBProject.VBComponents
        .Remove .item(moduleName)
    End With
    removemodule = Not (Err.Number = 9)
End Function

Private Function FromBase64(Text As String) As Byte()
    Dim Out() As Byte
    Dim b64(0 To 255) As Byte, str() As Byte, i&, j&, v&, b0&, b1&, b2&, b3&
    Out = ""
    If Len(Text) Then Else Exit Function

    str = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    For i = 2 To UBound(str) Step 2
        b64(str(i)) = i \ 2
    Next

    ReDim Out(0 To ((Len(Text) + 3) \ 4) * 3 - 1)
    str = Text & String$(2, 0)

    For i = 0 To UBound(str) - 7 Step 2
        b0 = b64(str(i))

        If b0 Then
            b1 = b64(str(i + 2))
            b2 = b64(str(i + 4))
            b3 = b64(str(i + 6))
            v = b0 * 262144 + b1 * 4096& + b2 * 64& + b3 - 266305
            Out(j) = v \ 65536
            Out(j + 1) = (v \ 256&) Mod 256
            Out(j + 2) = v Mod 256
            j = j + 3
            i = i + 6
        End If
    Next

    If b2 = 0 Then
        Out(j - 3) = (v + 65) \ 65536
        j = j - 2
    ElseIf b3 = 0 Then
        Out(j - 3) = (v + 1) \ 65536
        Out(j - 2) = ((v + 1) \ 256&) Mod 256
        j = j - 1
    End If

    ReDim Preserve Out(j - 1)
    FromBase64 = Out
End Function
