using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

using Microsoft.Extensions.Logging;

namespace Tlabs.CalcNgn.Parser {

  /// <summary>
  /// CommandTokens beinhaltet alle Begriffe nach dem der Parser sucht.
  /// </summary>
  public enum CommandTokens : int {
    /// <summary>@Agent</summary>
    Agent,
    /// <summary>@AgentTemplateRef</summary>
    AgentTemplateRef,
    /// <summary>@Trigger</summary>
    Trigger,
    /// <summary>@TriggerTemplateRef</summary>
    TriggerTemplateRef,
    /// <summary>@Param</summary>
    Param,
    /// <summary>@Scope</summary>
    Scope,
    /// <summary>@Readonly</summary>
    Readonly,
    /// <summary>@Assistent</summary>
    Assistent,
    /// <summary>@Help</summary>
    Help
  };

  /// <summary>
  /// Mögliche Parameter für die Begriffe die unter 
  ///  CommandTokens.Agent 
  ///  CommandTokens.Trigger 
  ///  CommandTokens.AgentTemplateRef 
  ///  CommandTokens.TriggerTemplateRef 
  /// definiert wurden.
  /// </summary>
  public enum TemplateParamTokens : int {
    /// <summary>ID</summary>
    ID=100,
    /// <summary>DESC</summary>
    DESC=101,
    /// <summary>CLASS</summary>
    CLASS=102,
    /// <summary>PARAM</summary>
    PARAM=103,
    /// <summary>AGENT</summary>
    AGENT=104,
    /// <summary>INFO</summary>
    INFO=105
  };

  /// <summary>
  /// Mögliche Parameter für den CommandTokens.Param
  /// </summary>
  public enum DatatypeTokens : int {
    /// <summary>STRING</summary>
    STRING=200,
    /// <summary>INT</summary>
    INT=201,
    /// <summary>FLOAT</summary>
    FLOAT=202,
    /// <summary>DATE</summary>
    DATE=203,
    /// <summary>MULTI</summary>
    MULTI=204
  };

  /// <summary>
  /// Mögliche Parameter für den CommandTokens.Scope
  /// </summary>
  public enum ScopeTokens : int {
    /// <summary>SYSTEM</summary>
    SYSTEM=300,
    /// <summary>USER</summary>
    USER=301
  };

  /// <summary>
  /// Mögliche Parameter für den CommandTokens.Readonly
  /// </summary>
  public enum ReadonlyToken : int {
    /// <summary>FALSE</summary>
    FALSE=400,
    /// <summary>TRUE</summary>
    TRUE=401
  };

  /// <summary>
  /// Durchsucht einen String nach den unter <see cref="CommandTokens">CommandTokens</see> angegebenen Begriffen.
  /// </summary>
  public class CmdTokenizer {

    /// <summary>
    /// Präfix für einen Begriff.
    /// Die unter CommandTokens definierten Begriffe werden nur dann gefunden wenn diese mit diesem Präfix versehen sind.
    /// </summary>
    public const String COMMAND_PREFIX = "@";

    readonly String[] aCommands;
    private const char ELEM_BRACKET_OPEN ='(';
    private const char ELEM_BRACKET_CLOSE =')';
    private const char ELEM_DELIMITER =',';
    private const char ELEM_DOUBLE_QUOTE ='"';
    private const char ELEM_ESCAPE ='\\';

    readonly String code;
    readonly IDictionary<String, IList<string>> dicCommands= new Dictionary<String, IList<string>>(StringComparer.InvariantCultureIgnoreCase);
    internal static readonly ILogger log= App.Logger<CmdTokenizer>();

    /// <summary>
    /// Durchsucht den als Parameter übergebenen Text nach Kommandos
    /// </summary>
    /// <remarks>
    /// Durchsucht den als Paramter übergebenen Text nach Kommandos (COMMANDS)
    /// Ein Kommando besteht immer aus einem Prefix(@) und dem Kommandonamen
    /// Bsp für AgentTemplate -> @AgentTemplate
    /// </remarks>
    /// <param name="text">String</param>
    public CmdTokenizer(String text) : this(typeof(CommandTokens), text) {}

    /// <summary>
    /// Durchsucht den als Parameter übergebenen Text nach Kommandos
    /// </summary>
    /// <remarks>
    /// Durchsucht den als Paramter übergebenen Text nach Kommandos (COMMANDS)
    /// Ein Kommando besteht immer aus einem Prefix(@) und dem Kommandonamen
    /// Bsp für AgentTemplate -> @AgentTemplate
    /// </remarks>
    /// <param name="CmdTokens">Type of Enum</param>
    /// <param name="text">String</param>
    public CmdTokenizer(Type CmdTokens, String text)
    {
      aCommands = Enum.GetNames(CmdTokens);
      code = text;
      searchAtCommands();
      //make read-only:
      // this.dicCommands = new Util.DictionaryExtension.ReadOnlyDictionary<string, IList<string>>(this.dicCommands);
    }

    /// <summary>
    /// Gibt ein Dictionary mit allen gefunden Kommandos und dessen Parameter zurück.
    /// </summary>
    /// <remarks>
    /// Gibt ein Dictionary mit allen gefunden Kommandos und dessen Parameter zurück.
    /// Die Kommandos werden ohne Prefix als Key im Dictionary hinterlegt.
    /// </remarks>
    public IDictionary<String, IList<string>> CommandDictionary {
      get { return dicCommands; }
    }

    /// <summary>
    /// Erstellt auf Basis der Enum Typen ein Kommando
    /// </summary>
    /// <param name="cmdToken"><see cref="CommandTokens"/></param>
    /// <returns></returns>
    public static String GetCommand(Enum cmdToken)
    {
      return GetCommand(cmdToken, null);
    }

    /// <summary>
    /// Gibt den Namen eines CommandTokens zurück.
    /// </summary>
    /// <param name="eCommand">Enum</param>
    /// <returns>String</returns>
    public static String GetCommandTokensName(Enum eCommand){
      return Enum.GetName(eCommand.GetType(), eCommand);
    }

    /// <summary>
    /// Erstellt auf Basis der Enum Typen ein Kommando
    /// </summary>
    /// <param name="cmdToken">Enum</param>
    /// <param name="Params">Siehe TemplateParamTokens, DatatypeTokens, ScopeTokens, ReadonlyToken </param>
    /// <returns></returns>
    public static String GetCommand(Enum cmdToken, params Enum[] Params) {
      return GetCommand(cmdToken, false, Params);
    }

    /// <summary>
    /// Erstellt auf Basis der Enum Typen ein Kommando
    /// </summary>
    /// <param name="cmdToken">Enum</param>
    /// <param name="woprefix">Commando wird ohne Prefix erstellt</param>
    /// <param name="Params">Siehe TemplateParamTokens, DatatypeTokens, ScopeTokens, ReadonlyToken </param>
    /// <returns></returns>
    public static String GetCommand(Enum cmdToken, bool woprefix, params Enum[] Params) {
      String sCommand= GetCommandTokensName(cmdToken);
      StringBuilder sbParam = new StringBuilder((woprefix ? "" : COMMAND_PREFIX)).Append(sCommand).Append('(');
      if (Params != null) {
        foreach (object param in Params)
          sbParam.Append(Enum.GetName(param.GetType(), param)).Append(',');
        if (sbParam[sbParam.Length - 1] == ',') sbParam.Length -= 1;
      }
      return sbParam.Append(')').ToString();
    }

    private void searchAtCommands(){

      int offset= 0;

      for (int i= 0; (i < aCommands.Length && offset < code.Length); i++) {
        offset= code.IndexOf(COMMAND_PREFIX, offset, StringComparison.Ordinal);
        if (offset == -1) break;
        offset= parseCommand(offset + COMMAND_PREFIX.Length);
      }

    }

    private int parseCommand(int offset) {

      for (int i= offset; i < code.Length; i++) {
        if (code[i] != ELEM_BRACKET_OPEN) continue;
        
        String sCommand= code.Substring(offset, i - offset).Trim();
        if (aCommands.Contains<String>(sCommand, StringComparer.InvariantCultureIgnoreCase)) {
          var lParam= new List<string>();
          dicCommands.Add(sCommand, new ReadOnlyCollection<string>(lParam));
          i= parseParam(i + 1, lParam);
        }
        else log.LogDebug("Unknown directive => {cmd}", sCommand);
        offset= i;
        break;
      }
      return offset;
    }

    private int parseParam(int offset, IList<string> lParam) {

      int lastMatch= offset;
      bool bCloseBracketFound= false;

      for (; offset < code.Length; offset++) {
        char key= code[offset];
        if (key == ELEM_DOUBLE_QUOTE) {
          int j= parseCloseDoubleQuote(++offset);
          lParam.Add(code.Substring(offset, j-offset));
          lastMatch= (offset= parseDelimiter(j)) + 1;
        }
        else if (key == ELEM_DELIMITER) {
          lParam.Add(code.Substring(lastMatch, offset-lastMatch).Trim());
          lastMatch= offset + 1;
        }
        else if (key == ELEM_BRACKET_CLOSE) {
          if (lastMatch+1 < offset) {
            lParam.Add(code.Substring(lastMatch, offset-lastMatch));
          }
          bCloseBracketFound= true;
          break;
        }
      }
      if (!bCloseBracketFound) {
        log.LogWarning("No closing bracket.");
      }
      return offset;
    }


    private int parseCloseDoubleQuote(int offset) {
      for (int i= offset; i < code.Length; i++) {
        if (code[i] == ELEM_DOUBLE_QUOTE && (i > 0 && code[i-1] != ELEM_ESCAPE)) return i;
      }
      return code.Length;
    }

    private int parseDelimiter(int offset) {
      for (int i= offset; i < code.Length; i++) {
        if (code[i] == ELEM_DELIMITER) return i;
        if (code[i] == ELEM_BRACKET_CLOSE) return i - 1;
      }
      return code.Length;
    }

  }
}
