/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.schneider.tsm.cprnprint;

import java.io.IOException;

/**
 *
 * @author SESA385697 Carlos Gonzalez Alonso @ Schneider-Electric, MDIC, TSM
 * Team.
 */
public class testss {
     public void sendMail(String ScriptToSent) throws InterruptedException
    {
        try
        {
            Process p = Runtime.getRuntime().exec("cmd.exe "+ "C:\\scripting\\" + ScriptToSent );
            System.out.println("cmd.exe /C "+ "C:\\scripting\\" + ScriptToSent);
            System.out.println("Mail Sended");
           // Thread.sleep(5000);
        }
        catch (IOException ioe)
        {
            System.out.println (ioe);
        }
    }
}
