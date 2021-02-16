Zapsi Service Terminal Special

LOOP EVERY 10 SECONDS
- program checks for open order
    - if open order has mode SERIZENI, program checks for divison
        - if division == 2, program checks for actual state
            - if production
                - program closes old terminal_input_order and start new order terminal_input_order
                - program sends PRODUCTION XML for main user and PRODUCTION XML for all additional users
            - if not production, program does nothing
        - if division == 3, program checks for actual state
            - if production for more than 15 minutes
                - program closes old terminal_input_order and start new order terminal_input_order
                - program sends PRODUCTION XML for main user and PRODUCTION XML for all additional users
            - if not production, program does nothing
    - if open order has not mode SERIZENI, or no open order, program does nothing
    
- program checks actual time
    - if time is in interval "15 minutes before shift ends" program checks for open order
        - if there is open order
            - program closes terminal_input_order and terminal_input_login
            - program sends TECHNOLOGY XML for main user, ENDWORK XML for main user, ENDWORK XML for all additional users and FINISH XML for main user
        - if there is not open order
            - program closes terminal_input_login
    - if time is not in interval, program does nothing
    
   