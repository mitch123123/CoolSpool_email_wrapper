      //-- ---------------------------------------------------------------------------
      // Application: send emails using coolspool by calling with PGM name
      //              uses femailer to find parms for coolspool command
      // Object Name: pemailer
      // Object Type: RPG SQL Bound Program
      // Object Date: 07/02/24
      //
      // Change History:
      //  mjf:creation
      //-----------------------------------------------------------------------------
       Ctl-Opt AlwNull(*USRCTL)  DatFmt(*Iso) TimFmt(*Iso) Debug(*Yes)
        Dftactgrp(*NO)  DecEdit('0.')
        Option(*Nodebugio:*SrcStmt:*NounRef);
       //////////////////////////////////////////////////////////
       //files
       //////////////////////////////////////////////////////////
       DCL-F fcstring keyed  USAGE(*OUTPUT) ; //log errors
       //////////////////////////////////////////////////////////
       //variables
       //////////////////////////////////////////////////////////
       dcl-s Message char(500);
       dcl-s count zoned(5:0) inz;
       dcl-s #ofParms packed(2:0);
       dcl-s #jobnam char(10);
       dcl-s #UserName char(10);
      //////////////////////////////////////////////////////////
      // external procedures
      //////////////////////////////////////////////////////////
       Dcl-Pr ExeClCmd EXTPGM('QCMDEXC');
         CmdStr Char(5000) Const Options(*VarSize);
         Len packed(15:5) Const;
       End-Pr;
      //////////////////////////////////////////////////////////
      // Parameters
      //////////////////////////////////////////////////////////
       dcl-pi *n;
        fpgm     char(10);
        fversion char(10);
       end-pi;
       /////////////////////////////////////////////////////////////
       //main
       /////////////////////////////////////////////////////////////
       EmailExcel(fpgm : fversion);

       // End of Pgm
       *inlr = *on;
        Return;
       ////////////////////////////////////////////////////////////
       // INIT
       ////////////////////////////////////////////////////////////
       begsr *INZSR;
          //delete sql history records over 8 days old
          EXEC SQL
          delete from fcstring
          where strhpgm = :fpgm
          and strhts < now() - 30 days ;
       endsr;
       //////////////////////////////////////////////////////////////
       //email the excel file out
       /////////////////////////////////////////////////////////////
       Dcl-Proc  EmailExcel;
        Dcl-Pi *N ;
          ppgm     char(10);
          pversion char(10);
        End-Pi ;
        DCL-DS Emailstuff extname('FEMAILER') inz END-DS;
        dcl-s commandtype   Char(10);
        dcl-s filetype   Char(10);
        dcl-s fileoption char(300);
        dcl-s FileEnding char(10);
        dcl-s Streamfile char(100);
        dcl-s overlay char(50);
        dcl-s filename char(10);
        dcl-s emailstr char(100);
        DCL-S EMAILOPTIONS CHAR(150);
        dcl-s sql char(100);
        dcl-s CMD   Char(5000);
        dcl-s ExcelHeader   Char(200);
        dcl-s Subject   Char(200);
        dcl-s Msg   Char(300);
        dcl-s #c_quote  char(1)   INZ(X'7D') ;
        Clear Cmd;
        exsr gatherrequirements; //retrieve femailer parameters
        exsr SetEmail; //check if email list or email
        exsr setConvertType;//set convert to command
        exsr setFileType;//set from file
        exsr setStrmf;//add datetime to name for unqiue
        exsr setFileNameANDoption; //add datetime to name for unqiue
        exsr setSubject;//add date to subject
        exsr setMessage;//add date and create program to email message

          IF FILETYPE = 'SPOOL';
             EXSR CraftSpoolCommand;
          else;
             exsr CraftCommand;
          ENDIF;
          exsr qcmdexecutestmt;
       //////////////////////////////////////////////////////////////
       //gather softcoded parameters for the command
       /////////////////////////////////////////////////////////////
         begsr gatherrequirements;
                EXEC SQL
                select * into :emailstuff
                from femailer
                where EID = :ppgm and EVERS = :pversion;
                if sqlcode >0;
                 //logerr
                endif;
         ENDSR;
       //////////////////////////////////////////////////////////////
       //set the email recipient
       /////////////////////////////////////////////////////////////
         begsr SetEmail;
            if %scan('@':eemail )>0; //check if email or addrlist
                 emailstr= %trim(eemail) +' *PRI';
            else;
                 emailstr= %trim(eemail) +' *ADRL *ADRL';
            ENDIF;
         ENDSR;
         BEGSR setEmailOptions;
           emailOptions= '';
         ENDSR;
       //////////////////////////////////////////////////////////////
       //what are we converting the file into,
       /////////////////////////////////////////////////////////////
         begsr setConvertType;
         if %TRIM(eftype) ='SPOOL';
            if ectype = 'EXCEL';
               commandtype='CVTSPLXLS';
               FileEnding ='.xls';
            ELSEIF ectype = 'PDF' ;
               commandtype='CVTSPLPDF';
               FileEnding ='.pdf';
            ELSEIF ectype = 'CSV' ;
               commandtype='CVTSPLCSV';
               FileEnding ='.csv';
            ELSEIF ectype = 'HTML' ;
               commandtype='CVTSPLHTML';
               FileEnding ='.html';
            ELSE;
              //  LOGERR('INVALID CONVERT TYPE');
            ENDIF;
          ELSE;
            if ectype = 'EXCEL';
               commandtype='CVTDBFXLSX';
               FileEnding ='.xlsx';
            ELSEIF ectype = 'CSV' ;
               commandtype='CVTDBFCSV';
               FileEnding ='.csv';
            ELSEIF ectype = 'HTML' ;
               commandtype='CVTDBFHTML';
               FileEnding ='.html';
            ELSE;
             // LOGERR('INVALID CONVERT TYPE');
            ENDIF;
         ENDIF;
         ENDSR;
       //////////////////////////////////////////////////////////////
       //VALID FILE TYPES:SQL,SQLSRC, DBF, SPOOL
       //this would be the from file
       /////////////////////////////////////////////////////////////
       BEGSR setFileType;
        filetype= EFTYPE;
       ENDSR;
       //////////////////////////////////////////////////////////////
       //set the stream file path
       /////////////////////////////////////////////////////////////
       BEGSR setStrmf;
        IF Estmf <>' ';
          Streamfile = Estmf;
        else;
          Streamfile ='/cool/'+%trim(EFILE)+
                      '.'+%trim(EFTYPE);
        ENDIF;
       ENDSR;
       //////////////////////////////////////////////////////////////
       //set the email message
       /////////////////////////////////////////////////////////////
       BEGSR setMessage;
          IF emessage <>' ';
           Msg = %trim(emessage) + ' Created by pgm: '+%trim(ppgm)+
                   ' On date: ' + %Char(%Date():*Usa) +
                   ' At time: '+%Char(%time():*Usa );
          else;
           Msg='need message';
          ENDIF;
       ENDSR;
       //////////////////////////////////////////////////////////////
       //set the email subject
       /////////////////////////////////////////////////////////////
       begsr setSubject;
         IF esubject <>' ';
           subject =%trim(esubject) + ' '+%Char(%Date():*Usa);
         else;
           subject = 'need subject';
         ENDIF;
       ENDSR;
       begsr setOverlay;
         IF EOVERLAY <> ' ';
             overlay     = ' INCLFILE(('+%trim(EOVERLAY)+
                           ' *JPG *EMBEDDED *ALL 0 0 '+
                           ' *MM *NONE *NONE 1)) ';
         ENDIF;
       ENDSR;

       //////////////////////////////////////////////////////////////
       //set file name and potential options
       /////////////////////////////////////////////////////////////
       begsr setFileNameANDoption;
         if eftype ='SQL';
           filename = '*SQL';
           fileoption = ' SQL('''
           +%trim(esql) +''') ';

         ELSEIF eftype ='SQLSRC';
            filename = '*SQLSRC';
            fileoption = ' SQLSRC('+%trim(ELIB)+'/'+%trim(EFILE)
            +' '+%trim(EMEMBER)+') ';
         ELSE;
            Filename = %trim(ELIB) +'/'+%trim(EFILE);
            fileoption = ' ';
         ENDIF;
       endsr;
       //////////////////////////////////////////////////////////////
       //email the excel file out
       /////////////////////////////////////////////////////////////
       BEGSR CraftCommand;
           cmd = commandtype +
                 ' FROMFILE('+%trim(filename) +') ' +
                 ' TOSTMF(''' + %trim(Streamfile) +''') ' +
                 ' STMFOPT(*REPLACE) ' +
                 %trim(fileoption) + //for sql, sqlsrc, otherwise blank
                 ' EMAIL(*YES) '+
                 ' EMAILOPT(*NO ''' + %trim(Subject) +''' ) '+
                 ' EMAILFROM(DoNotReply@YOURDOMAIN.com) ' +
                 ' EMAILTO(('+%trim(emailstr)+')) ' +
                 ' EMAILMSG(''' + %trim(Msg) + ''')' ;
       ENDSR;
       //////////////////////////////////////////////////////////////
       //email the sppol file out
       /////////////////////////////////////////////////////////////
       BEGSR CraftSpoolCommand;
           cmd = commandtype +
                 ' FROMFILE('+%trim(filetype) +') ' +
                 ' TOSTMF(''' + %trim(Streamfile) +''') ' +
                 ' STMFOPT(*REPLACE) ' +
                 %trim(fileoption) + //for sql, sqlsrc, otherwise blank
                 ' EMAIL(*YES) '+
                 ' EMAILOPT(*NO ''' + %trim(Subject) +''' ) '+
                 ' EMAILFROM(DoNotReply@YOURDOMAIN.com) ' +
                 ' EMAILTO(('+%trim(emailstr)+')) ' +
                 ' EMAILMSG(''' + %trim(Msg) + ''')' ;
       ENDSR;
       //////////////////////////////////////////////////////////////
       //log error
       /////////////////////////////////////////////////////////////
         begsr qcmdexecutestmt;

          monitor;
            ExeClCmd(cmd : %Len(%Trim(cmd)));
           on-error;
              clear rcstring;
              strhPGM  = ppgm;
              strhID   = pversion;
              strhTS   = %Timestamp();
              strhUSER = #UserName;
              strhTYPE = 'CMD';
              strhCODE = -9999;
              strhSTMT = %trim(cmd);

              write rcstring;


              return;
          endmon;


       endsr;
       end-proc;
