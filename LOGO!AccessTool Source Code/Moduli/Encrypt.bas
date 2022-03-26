Attribute VB_Name = "Encrypt"
Option Private Module
Dim m_crc32table(255)

Public Sub AppendCrc32Table(offset, data)
    Dim i
    For i = LBound(data) To UBound(data) Step 1
        m_crc32table(offset + i) = data(i)
    Next
End Sub

Public Function MakeCRC32(uint8arr)
    If m_crc32table(1) <> 1996959894# Then
        AppendCrc32Table 0, Array(0#, 1996959894#, 3993919788#, 2567524794#, 124634137#, 1886057615#, 3915621685#, 2657392035#, 249268274#, 2044508324#, 3772115230#, 2547177864#, 162941995#, 2125561021#, 3887607047#, 2428444049#)
        AppendCrc32Table 16, Array(498536548#, 1789927666#, 4089016648#, 2227061214#, 450548861#, 1843258603#, 4107580753#, 2211677639#, 325883990#, 1684777152#, 4251122042#, 2321926636#, 335633487#, 1661365465#, 4195302755#, 2366115317#)
        AppendCrc32Table 32, Array(997073096#, 1281953886#, 3579855332#, 2724688242#, 1006888145#, 1258607687#, 3524101629#, 2768942443#, 901097722#, 1119000684#, 3686517206#, 2898065728#, 853044451#, 1172266101#, 3705015759#, 2882616665#)
        AppendCrc32Table 48, Array(651767980#, 1373503546#, 3369554304#, 3218104598#, 565507253#, 1454621731#, 3485111705#, 3099436303#, 671266974#, 1594198024#, 3322730930#, 2970347812#, 795835527#, 1483230225#, 3244367275#, 3060149565#)
        AppendCrc32Table 64, Array(1994146192#, 31158534#, 2563907772#, 4023717930#, 1907459465#, 112637215#, 2680153253#, 3904427059#, 2013776290#, 251722036#, 2517215374#, 3775830040#, 2137656763#, 141376813#, 2439277719#, 3865271297#)
        AppendCrc32Table 80, Array(1802195444#, 476864866#, 2238001368#, 4066508878#, 1812370925#, 453092731#, 2181625025#, 4111451223#, 1706088902#, 314042704#, 2344532202#, 4240017532#, 1658658271#, 366619977#, 2362670323#, 4224994405#)
        AppendCrc32Table 96, Array(1303535960#, 984961486#, 2747007092#, 3569037538#, 1256170817#, 1037604311#, 2765210733#, 3554079995#, 1131014506#, 879679996#, 2909243462#, 3663771856#, 1141124467#, 855842277#, 2852801631#, 3708648649#)
        AppendCrc32Table 112, Array(1342533948#, 654459306#, 3188396048#, 3373015174#, 1466479909#, 544179635#, 3110523913#, 3462522015#, 1591671054#, 702138776#, 2966460450#, 3352799412#, 1504918807#, 783551873#, 3082640443#, 3233442989#)
        AppendCrc32Table 128, Array(3988292384#, 2596254646#, 62317068#, 1957810842#, 3939845945#, 2647816111#, 81470997#, 1943803523#, 3814918930#, 2489596804#, 225274430#, 2053790376#, 3826175755#, 2466906013#, 167816743#, 2097651377#)
        AppendCrc32Table 144, Array(4027552580#, 2265490386#, 503444072#, 1762050814#, 4150417245#, 2154129355#, 426522225#, 1852507879#, 4275313526#, 2312317920#, 282753626#, 1742555852#, 4189708143#, 2394877945#, 397917763#, 1622183637#)
        AppendCrc32Table 160, Array(3604390888#, 2714866558#, 953729732#, 1340076626#, 3518719985#, 2797360999#, 1068828381#, 1219638859#, 3624741850#, 2936675148#, 906185462#, 1090812512#, 3747672003#, 2825379669#, 829329135#, 1181335161#)
        AppendCrc32Table 176, Array(3412177804#, 3160834842#, 628085408#, 1382605366#, 3423369109#, 3138078467#, 570562233#, 1426400815#, 3317316542#, 2998733608#, 733239954#, 1555261956#, 3268935591#, 3050360625#, 752459403#, 1541320221#)
        AppendCrc32Table 192, Array(2607071920#, 3965973030#, 1969922972#, 40735498#, 2617837225#, 3943577151#, 1913087877#, 83908371#, 2512341634#, 3803740692#, 2075208622#, 213261112#, 2463272603#, 3855990285#, 2094854071#, 198958881#)
        AppendCrc32Table 208, Array(2262029012#, 4057260610#, 1759359992#, 534414190#, 2176718541#, 4139329115#, 1873836001#, 414664567#, 2282248934#, 4279200368#, 1711684554#, 285281116#, 2405801727#, 4167216745#, 1634467795#, 376229701#)
        AppendCrc32Table 224, Array(2685067896#, 3608007406#, 1308918612#, 956543938#, 2808555105#, 3495958263#, 1231636301#, 1047427035#, 2932959818#, 3654703836#, 1088359270#, 936918000#, 2847714899#, 3736837829#, 1202900863#, 817233897#)
        AppendCrc32Table 240, Array(3183342108#, 3401237130#, 1404277552#, 615818150#, 3134207493#, 3453421203#, 1423857449#, 601450431#, 3009837614#, 3294710456#, 1567103746#, 711928724#, 3020668471#, 3272380065#, 1510334235#, 755167117#)
    End If
    
    Dim u32CRC
    u32CRC = 4294967295#
    
    Dim u8TableUnitLocation
    
    Dim i
    For i = LBound(uint8arr) To UBound(uint8arr) Step 1
        u8TableUnitLocation = GetLowerBits(u32CRC, 8) Xor uint8arr(i)
        u32CRC = CalculateXOR(MoveRight(u32CRC, 8), m_crc32table(u8TableUnitLocation))
    Next
    
    MakeCRC32 = CalculateXOR(u32CRC, 4294967295#)
End Function


Function GetLowerBits(val, bits)
    Dim factor
    factor = (2 ^ bits)
    GetLowerBits = val - Int(val / factor) * factor
End Function

Function MoveRight(val, bits)
    MoveRight = Int(val / (2 ^ bits))
End Function

Function CalculateXOR(v1, v2)
    CalculateXOR = (MoveRight(v1, 8) Xor MoveRight(v2, 8)) * 256 + (GetLowerBits(v1, 8) Xor GetLowerBits(v2, 8))
End Function

Function TestBit(val, bits)
    TestBit = (val \ (2 ^ bits)) Mod 2
End Function

Function String2UTF8(STR)
    Dim data()
    Dim Count

    ReDim data(256) ' initial dim of the array
    Count = 0
    
    Dim i
    Dim charCount
    
    charCount = Len(STR)
    For i = 1 To charCount Step 1
        Dim code
        code = AscW(Mid(STR, i, 1))
        
        ' format unicode code of the character i
        If code < 0 Then
            code = code + 65536
        End If
        
        ' ensure there are at least 4 vacants
        If UBound(data) < Count + 4 Then
            ReDim Preserve data(Count + 256) ' 256 for each time
        End If
        
        ' fill code into data array
        If code < &H80 Then
            data(Count) = code
            Count = Count + 1
        Else
            If code < &H800 Then
                ' 2 bytes char:
                data(Count) = (MoveRight(code, 6) And 31) Or 192
                Count = Count + 1
                data(Count) = (code And 63) Or 128
                Count = Count + 1
            Else
                If code < &H10000 Then
                    ' 3 bytes char:
                    data(Count) = (MoveRight(code, 12) And 15) Or 224
                    Count = Count + 1
                    data(Count) = (MoveRight(code, 6) And 63) Or 128
                    Count = Count + 1
                    data(Count) = (code And 63) Or 128
                    Count = Count + 1
                Else
                    If code < &H110000 Then
                        ' 4 bytes char:
                        data(Count) = (MoveRight(code, 18) And 7) Or 240
                        Count = Count + 1
                        data(Count) = (MoveRight(code, 12) And 63) Or 128
                        Count = Count + 1
                        data(Count) = (MoveRight(code, 6) And 63) Or 128
                        Count = Count + 1
                        data(Count) = (code And 63) Or 128
                        Count = Count + 1
                    Else
                        ' invalid uncode, ignored.
                    End If
                End If
            End If
        End If
    Next
    
    ' format the data before returning
    ReDim Preserve data(Count - 1)
    
    'MsgBox UBound(data)
    
    String2UTF8 = data
End Function
