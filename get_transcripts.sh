#!/usr/bin/env bash

API_KEY=""
mkdir -p transcripts

i=1
while IFS= read -r url; do
  [ -z "$url" ] && continue

  out=$(printf "transcripts/%04d.txt" "$i")

  curl -sS -G "https://transcriptapi.com/api/v2/youtube/transcript" \
    --data-urlencode "video_url=$url" \
    --data-urlencode "format=text" \
    --data-urlencode "include_timestamp=false" \
    -H "Authorization: Bearer $API_KEY" \
    -o "$out"

  echo "Saved $url -> $out"
  i=$((i + 1))
  sleep 0.5
done < youtube_urls.txt
