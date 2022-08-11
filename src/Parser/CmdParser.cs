using System;
using System.Collections;
using System.Collections.Generic;

namespace Tlabs.CalcNgn.Parser {

  /// <summary>
  /// Eventhandler
  /// </summary>
  /// <param name="Sender">Object</param>
  /// <param name="Command">String - Name des gefunden Begriffs</param>
  /// <param name="dicParserResult"><see cref="IDictionary"/></param>
  /// <param name="column">int</param>
  /// <param name="row">int</param>
  public delegate void CommandFoundEventHandler(object Sender, String Command, IDictionary<String, IList<string>> dicParserResult, int column, int row);
  
  /// <summary>
  /// Verwendet den <see cref="CmdTokenizer"/> um ein String nach registrierten Begriffen zu dursuchen und
  /// löst ein <see cref="CommandFoundEventHandler">Event</see> aus wenn dieser Begriff in dem String gefunden wurde.
  /// </summary>
  public class CmdParser {
    readonly Type cmdEnumType;
    private Dictionary<String, RegEvent> dicRegEV;

    ///<summary>Default ctor</summary>
    public CmdParser() : this(typeof(CommandTokens)) { }

    ///<summary>Ctor from <paramref name="cmdEnumType"/></summary>
    public CmdParser(Type cmdEnumType) {
      if (null == (this.cmdEnumType= cmdEnumType)) throw new ArgumentNullException(nameof(cmdEnumType));
      init();
    }

    private void init() {
      dicRegEV= new Dictionary<String, RegEvent>();
    }

    ///<summary>
    ///Setzt den Interpreter zurück.
    ///D.h. alle Event - Registrierung werden gelöscht.
    ///</summary>
    public void reset() {
      init();
    }
    
    ///<summary>
    ///Analysiert den als Parameter übergebenen Text.
    ///</summary>
    ///<remarks>
    ///Analysiert den als Parameter übergebenen Text.
    ///Die Paramter column und row werden als Information an den Event weitergegeben 
    ///</remarks>
    ///<param name="txt">txt</param>
    ///<param name="column">int</param>
    ///<param name="row">int</param>
    public void parse(String txt, int column, int row){
      var p= new CmdTokenizer(cmdEnumType, txt);
      var dicResult= p.CommandDictionary;

      foreach (KeyValuePair<String, RegEvent> kvp in dicRegEV) {
        if (!dicResult.TryGetValue(kvp.Key, out var lParam)) continue;
        if (kvp.Value.paramList.Count == 0 || kvp.Value.RunOnce) {
          kvp.Value.eventHandler(this, kvp.Key, dicResult, column, row);
        }
        else if (lParam.Count > 0) {
          foreach (String param in lParam) {
            if (kvp.Value.paramList.Count > 0 
                && !kvp.Value.paramList.Contains(param)) continue;
            kvp.Value.eventHandler(this, kvp.Key, dicResult, column, row);
          }
        }
      }

    }

    /// <summary>
    /// Registriert einen Event für ein Kommando
    /// </summary>
    /// <remarks>
    /// Registriert einen Event für ein Kommando
    /// Das Kommando muß mit Prefix übergeben werden.
    /// Paremeter werden für die Registrierung berücksichtigt.
    /// Bsp: @AgentInstance(ID,DESC)
    /// Der Event wird jetzt sowohl für ID als auch für DESC ausgelöst.
    /// </remarks>    
    /// <param name="p">Parser</param>
    /// <param name="ev">CommandFindEventHandler</param>
    /// <param name="runonce">Boolean - True Der event wird nur einmal aufgerufen und nicht wie sonst für jeden Parameter</param>
    public void registerEvent(CmdTokenizer p, CommandFoundEventHandler ev, Boolean runonce) {
      var dicResult= p.CommandDictionary;

      foreach (KeyValuePair<String, IList<string>> kvp in dicResult) {
        if (dicRegEV.ContainsKey(kvp.Key)) continue;
        var rev= new RegEvent(kvp.Value, ev, runonce);
        dicRegEV.Add(kvp.Key, rev);
      }
    }

    /// <summary>
    /// Deregistriert einen Event für ein Kommando 
    /// Das Kommando wird ohne Prefix übergeben.
    /// </summary>
    /// <param name="command">String</param>
    public void deregisterEvent(String command) {
      if (!dicRegEV.ContainsKey(command)) return;
      dicRegEV.Remove(command);
    }

    /// <summary>
    /// Gibt eine Liste mit allen Kommandos (ohne Prefix), für die ein Event registriert wurde, zurück.
    /// </summary>
    /// <returns></returns>
    public ICollection<string> registeredEvents() {
      return dicRegEV.Keys;
    }

  }

  ///<summary>
  /// Wird vom INterpreter für die Event - Registrierung verwendet.
  ///</summary>
  class RegEvent {

    readonly IList<string> lParam;
    readonly CommandFoundEventHandler ev;
    readonly Boolean runonce;

    public RegEvent(IList<string> lparam, CommandFoundEventHandler evc, Boolean runonce) {
      this.lParam= lparam;
      this.ev= evc;
      this.runonce= runonce;
    }

    public IList<string> paramList {
      get { return lParam; }
    }

    public CommandFoundEventHandler eventHandler {
      get { return ev; }
    }

    public Boolean RunOnce {
      get { return runonce; }
    }
  }
}
