$currentDirectory = $myinvocation.mycommand.Definition | split-path -parent
[Reflection.Assembly]::LoadFile("$currentDirectory\Twitterizer2.dll")
[Reflection.Assembly]::LoadFile("$currentDirectory\Newtonsoft.Json.dll")

Clear-Host

$t = Import-Csv -path "profile.csv" -Header "key", "value"
$profile = @{}
foreach($r in $t)
{
    $profile[$r.key] = $r.value
}

$tokens = New-Object Twitterizer.OAuthTokens
$tokens.AccessToken = $profile.AccessToken
$tokens.AccessTokenSecret = $profile.AccessTokenSecret
$tokens.ConsumerKey = $profile.ConsumerKey
$tokens.ConsumerSecret = $profile.ConsumerSecret

$timelineOptions = New-Object Twitterizer.timelineOptions
$timelineOptions.Count = 20
$timelineOptions.UseSSL = $true

$searchOptions = New-Object Twitterizer.SearchOptions
$searchOptions.NumberPerPage = 30
$searchOptions.UseSSL = $true

$retweetOptions = New-Object Twitterizer.OptionalProperties
$retweetOptions.UseSSL = $true

$statusUpdateOptions = New-Object Twitterizer.statusUpdateOptions
$statusUpdateOptions.UseSSL = $true

$command =  ""
$lastLatestStatusID = ""

While ($command -ne "Q")
{
    $statuses = @{}
    $pageNumber = $timelineOptions.Page
    $updateOrderNo = (($pageNumber-1)*$timelineOptions.Count+1)

    if ($command -eq "S" -or $command -eq "M")
    {}
    else {
        $response = [Twitterizer.TwitterTimeline]::HomeTimeline($tokens,$timelineOptions);
    }


    if ($response.Result -eq [Twitterizer.RequestResult]::Success)
    {
        clear-host
        Write-Host "- Page $pageNumber -" -ForegroundColor Green
        Write-Host

        foreach ($status in $response.ResponseObject)
        {
            if ($lastLatestStatusID -eq $status.Id)
            {
                Write-Host "---------- End of New Updates ----------" -ForegroundColor Magenta
                Write-Host
            }

            if ($status.User.Name)
            {
                Write-Host [$updateOrderNo] $status.User.Name -ForegroundColor Yellow
            }
            else 
            {
                Write-Host [$updateOrderNo] $status.FromUserScreenName -ForegroundColor Yellow
            }
            Write-Host $status.Text -NoNewline; Write-Host "("$status.CreatedDate")" -ForegroundColor Gray
            Write-Host

            $statuses.Add($updateOrderNo, $status)
            
            $updateOrderNo++
        }

        
        if ($pageNumber -eq 1)
        {
            $first = 1
            $lastLatestStatusID = $statuses.$first.Id
        }

        Write-Host "Options: (r)efresh, (t)weet, ret(w)eet, (n)ext page, (l)ast page, (m)entions, (o)pen link, (s)earch or (q)uit?" -ForegroundColor Green
        $command = [Console]::ReadKey($true).Key

        switch ($command)
        {
            "R" {
                $timelineOptions.Page = 1
            }

            "T" {
                $update = Read-Host "What do you want to say?(enter q!! to cancel the operation)"
                if ($update.ToLower() -ne "q!!")
                {
                    $updateResponse = [Twitterizer.TwitterStatus]::Update($tokens, $update, $statusUpdateOptions)
                    if ($updateResponse.Result -eq [Twitterizer.RequestResult]::Success)
                    {
                        Write-Host "Update Succeed!" -ForegroundColor Green
                    } else {
                        Write-Host "Update Failed!" -ForegroundColor Red
                        }
                } else {
                    Write-Host "Update canceled!"
                    }
            }

            "W" {
                [int]$retweetsNumber = Read-Host "Which tweet do you want to retweet?"
                $retweetResponse = [Twitterizer.TwitterStatus]::Retweet($tokens, $statuses.$retweetsNumber.ID, $retweetOptions)
                if ($retweetResponse.Result -eq [Twitterizer.RequestResult]::Success)
                {
                    Write-Host "Retweet Succeed!" -ForegroundColor Green
                } else {
                        Write-Host "Retweet Failed!" -ForegroundColor Red
                  }
                Read-Host
            }

            "N" {
                $timelineOptions.Page++
            }

            "L" {
                if ($timelineOptions.Page -eq 1)
                {
                    Write-Host "You are already at the first page" -ForegroundColor Red
                }
                else {
                    $timelineOptions.Page--
                }            
            }

            "O" {
                [int]$number = Read-Host "Which tweet?"
                foreach ($status in $statuses.$number)
                {
                    foreach ($entity in $status.Entities)
                    {
                        if ($entity.Url)
                        {
                            Start-Process $entity.Url
                        }
                    }
                }
            }

            "M" {
                $response = [Twitterizer.TwitterTimeline]::Mentions($tokens,$timelineOptions);
            }

            "S" {
                $searchTerm = Read-Host "What do you want to search?"
                $response = [Twitterizer.TwitterSearch]::search($tokens, $searchTerm, $searchOptions)
            }

            default { $timelineOptions.Page = 1 }
        }
        
    }

    else
    {
        Write-Host $response.ErrorMessage -ForegroundColor Red
        break
    }

}

Write-Host
Write-Host "Bye Bye! This script is powered by Twitterizer (http://www.twitterizer.net)."