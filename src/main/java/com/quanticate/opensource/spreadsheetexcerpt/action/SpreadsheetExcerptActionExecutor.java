package com.quanticate.opensource.spreadsheetexcerpt.action;

import java.util.List;

import org.alfresco.repo.action.executer.ActionExecuterAbstractBase;
import org.alfresco.service.cmr.action.Action;
import org.alfresco.service.cmr.action.ParameterDefinition;
import org.alfresco.service.cmr.repository.ContentService;
import org.alfresco.service.cmr.repository.NodeRef;
import org.alfresco.service.cmr.repository.NodeService;

import com.quanticate.opensource.spreadsheetexcerpt.excerpt.MakeReadOnlyAndExcerpt;

public class SpreadsheetExcerptActionExecutor extends ActionExecuterAbstractBase
{
   public static final String NAME = "spreadsheet-excerpt";
   
   private NodeService nodeService;
   private ContentService contentService;
   private MakeReadOnlyAndExcerpt excerpter;

   @Override
   protected void executeImpl(Action action, NodeRef nodeRef)
   {
      // TODO Action Definition stuff
      
      // Grab the details .....
   }

   @Override
   protected void addParameterDefinitions(List<ParameterDefinition> arg0)
   {
      // TODO Decide if we'll use this or not later
   }

   public void setNodeService(NodeService nodeService)
   {
      this.nodeService = nodeService;
   }
   public void setContentService(ContentService contentService)
   {
      this.contentService = contentService;
   }
   public void setExcerpter(MakeReadOnlyAndExcerpt excerpter)
   {
      this.excerpter = excerpter;
   }
}
