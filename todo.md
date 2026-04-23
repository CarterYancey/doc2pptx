[] Allow users to specify the generated ppt name

[] Large documents/transcripts should be chunked before given to the LLM to improve performance.

[] Separate the YT transcript-to-ppt-outline into a different project that relies on doc2pptx.

[] YT-to-outline should be condensed into one script/main-function that can be looped through a list, rather than expecting a list

[] get_transcripts.sh should save the transcript filename as the video name

[] get_transcripts.sh should fetch the API key from env variable or accepts as input parameter
