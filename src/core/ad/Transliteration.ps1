function Convert-UA2Latin {
    param([Parameter(Mandatory=$true)][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

    # Нормалізація пробілів + апострофів (ASCII-only)
    $t = $Text.Trim()
    $t = $t -replace "[`'’ʼ‘]", ""     # різні апострофи -> прибрати
    $t = $t -replace "\u00A0", " "     # NBSP -> space

    $sb = New-Object System.Text.StringBuilder
    $isWordStart = $true

    foreach ($ch in $t.ToCharArray()) {
        $code = [int][char]$ch

        # Latin letters / digits
        if (($code -ge 48 -and $code -le 57) -or ($code -ge 97 -and $code -le 122) -or ($code -ge 65 -and $code -le 90)) {
            [void]$sb.Append(([string]$ch).ToLower())
            $isWordStart = $false
            continue
        }

        # space or hyphen
        if ($code -eq 32 -or $code -eq 45) {
            [void]$sb.Append([char]$code)
            $isWordStart = $true
            continue
        }

        # To lower for Cyrillic ranges: we will handle both upper/lower by mapping both
        $out = $null

        # Word-start special rules (Ye/Yi/Y/Yu/Ya) in lowercase: ye/yi/y/yu/ya
        if ($isWordStart) {
            switch ($code) {
                0x0404 { $out = "ye" } # Є
                0x0454 { $out = "ye" } # є
                0x0407 { $out = "yi" } # Ї
                0x0457 { $out = "yi" } # ї
                0x0419 { $out = "y"  } # Й
                0x0439 { $out = "y"  } # й
                0x042E { $out = "yu" } # Ю
                0x044E { $out = "yu" } # ю
                0x042F { $out = "ya" } # Я
                0x044F { $out = "ya" } # я
            }
        }

        if (-not $out) {
            switch ($code) {
                # А а
                0x0410 { $out = "a" }; 0x0430 { $out = "a" }
                # Б б
                0x0411 { $out = "b" }; 0x0431 { $out = "b" }
                # В в
                0x0412 { $out = "v" }; 0x0432 { $out = "v" }
                # Г г (UA: h)
                0x0413 { $out = "h" }; 0x0433 { $out = "h" }
                # Ґ ґ (g)
                0x0490 { $out = "g" }; 0x0491 { $out = "g" }
                # Д д
                0x0414 { $out = "d" }; 0x0434 { $out = "d" }
                # Е е
                0x0415 { $out = "e" }; 0x0435 { $out = "e" }
                # Є є (not word start: ie)
                0x0404 { if (-not $out) { $out = "ie" } }
                0x0454 { if (-not $out) { $out = "ie" } }
                # Ж ж
                0x0416 { $out = "zh" }; 0x0436 { $out = "zh" }
                # З з
                0x0417 { $out = "z" }; 0x0437 { $out = "z" }
                # И и
                0x0418 { $out = "y" }; 0x0438 { $out = "y" }
                # І і
                0x0406 { $out = "i" }; 0x0456 { $out = "i" }
                # Ї ї (not word start: i)
                0x0407 { if (-not $out) { $out = "i" } }
                0x0457 { if (-not $out) { $out = "i" } }
                # Й й (not word start: i)
                0x0419 { if (-not $out) { $out = "i" } }
                0x0439 { if (-not $out) { $out = "i" } }
                # К к
                0x041A { $out = "k" }; 0x043A { $out = "k" }
                # Л л
                0x041B { $out = "l" }; 0x043B { $out = "l" }
                # М м
                0x041C { $out = "m" }; 0x043C { $out = "m" }
                # Н н
                0x041D { $out = "n" }; 0x043D { $out = "n" }
                # О о
                0x041E { $out = "o" }; 0x043E { $out = "o" }
                # П п
                0x041F { $out = "p" }; 0x043F { $out = "p" }
                # Р р
                0x0420 { $out = "r" }; 0x0440 { $out = "r" }
                # С с
                0x0421 { $out = "s" }; 0x0441 { $out = "s" }
                # Т т
                0x0422 { $out = "t" }; 0x0442 { $out = "t" }
                # У у
                0x0423 { $out = "u" }; 0x0443 { $out = "u" }
                # Ф ф
                0x0424 { $out = "f" }; 0x0444 { $out = "f" }
                # Х х
                0x0425 { $out = "kh" }; 0x0445 { $out = "kh" }
                # Ц ц
                0x0426 { $out = "ts" }; 0x0446 { $out = "ts" }
                # Ч ч
                0x0427 { $out = "ch" }; 0x0447 { $out = "ch" }
                # Ш ш
                0x0428 { $out = "sh" }; 0x0448 { $out = "sh" }
                # Щ щ
                0x0429 { $out = "shch" }; 0x0449 { $out = "shch" }
                # Ь ь (skip)
                0x042C { $out = "" }; 0x044C { $out = "" }
                # Ю ю (not word start: iu)
                0x042E { if (-not $out) { $out = "iu" } }
                0x044E { if (-not $out) { $out = "iu" } }
                # Я я (not word start: ia)
                0x042F { if (-not $out) { $out = "ia" } }
                0x044F { if (-not $out) { $out = "ia" } }
            }
        }

        if ($null -ne $out) {
            [void]$sb.Append($out)
            $isWordStart = $false
        }
        else {
            # інші символи ігноруємо
        }
    }

    return $sb.ToString()
}
