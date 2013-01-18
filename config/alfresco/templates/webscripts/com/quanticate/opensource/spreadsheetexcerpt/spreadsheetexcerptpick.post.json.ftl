<#escape x as jsonUtils.encodeJSONString(x)>
{[
  <#list sheets as sn>
    "${sn}"<#if sn_has_next>,</#if>
  </#list>
]}
</#escape>
