using System.Collections.Generic;
using UnityEngine;

public abstract class BaseScriptableObject<T> : ScriptableObject where T : BaseTable
{
    public List<T> rows = new();
}