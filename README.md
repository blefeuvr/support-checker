## support-checker

This script is an automatic tool to warn about time period without support assignee  
It requests Microsoft graph API for calendar then sends message through slack api  

- daily_notify:  
Check for today support assignees and send support planning in #support channel finally it calls weekly_notify  
daily_notify is meant to be called each morning from monday to friday  

- tomorrow_notify:  
Check fow tomorrow support assignees and send them personal msg reminder  
tomorrow_notify is meant to be called each afternoon from sunday to thursday  

- weekly_notify:  
Check for 15 next days not supported time and send warning to #support channel  
