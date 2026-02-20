@echo off
setlocal
cd /d "C:\PDM-SW\\SW_CACHE\\42e77f6a\\payload"
echo ==== START %DATE% %TIME% ====>>"C:\PDM-SW\\SW_CACHE\\42e77f6a\\payload\payload.log"
"C:\PDM-SW\\SW_CACHE\\42e77f6a\\payload\\PDM_SW_PAYLOAD.exe" --pdm-root "C:\PDM-SW" --workspace 42e77f6a --sw-context-file "C:\PDM-SW\\SW_CACHE\\42e77f6a\\payload\sw_context.json" --log-file "C:\PDM-SW\\SW_CACHE\\42e77f6a\\payload\payload.log" >> "C:\PDM-SW\\SW_CACHE\\42e77f6a\\payload\payload.log" 2>&1
echo ==== EXIT %ERRORLEVEL% %DATE% %TIME% ====>>"C:\PDM-SW\\SW_CACHE\\42e77f6a\\payload\payload.log"
endlocal
