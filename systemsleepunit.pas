unit SystemSleepUnit;

{$mode delphi}

interface

uses
 Windows;

type
  EXECUTION_STATE = DWORD;

const
  ES_SYSTEM_REQUIRED   = $00000001;  //Forces the system to be in the working state by resetting the system idle timer
  ES_DISPLAY_REQUIRED  = $00000002;  //Forces the display to be on by resetting the display idle timer
  ES_AWAYMODE_REQUIRED = $00000040;  //This value must be specified with ES_CONTINUOUS
  ES_CONTINUOUS        = $80000000;  //Informs the system that the state being set should remain in effect until the next call


//The thread must call SetThreadExecutionState periodically to prevent sleep mode
function SetThreadExecutionState(esFlags: EXECUTION_STATE): EXECUTION_STATE;
  stdcall; external 'kernel32.dll' name 'SetThreadExecutionState';

implementation

end.
