package com.quanticate.opensource.spreadsheetexcerpt.excerpt;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Set;

public interface MakeReadOnlyAndExcerpt
{
   /**
    * @return The Mimetype(s) that are supported
    */
   public Set<String> getSupportedMimeTypes();
   
   /**
    * @return The Sheetnames found in the given file
    */
   public String[] getSheetNames(File f) throws IOException;
   
   /**
    * Marks the sheets to keep as read only (removing formulas), and
    *  removes all other sheets
    */
   public void excerpt(String[] sheetsToKeep, File input, OutputStream output) throws IOException;
   
   /**
    * Marks the sheets to keep as read only (removing formulas), and
    *  removes all other sheets
    */
   public void excerpt(int[] sheetsToKeep, File input, OutputStream output) throws IOException;
}
