using System;
using System.Collections.Generic;

namespace BarnData.Core.Validation
{
    // Single source of truth for grade allow-lists and condemnation detection.
    //
    // Rationale: previously the cow/bull allow-lists were inlined in two places
    // (the Hot Weight import pipeline and the MarkKilledApi save endpoint). They
    // could drift apart, leading to rows that pass Hot Weight validation but get
    // rejected at Mark Killed save time. Centralizing them here removes that risk.
    //
    // Grades come in three classes:
    //   1. Production grades — animal was killed normally; sides/HS/HotWeight required.
    //      Cow:  CN, SH, CT, B1, B2, BR, NP
    //      Bull: BB, LB, UB, FB, NP
    //      (NP = Nicholas Prime — manually entered post-pull, valid for both types.)
    //
    //   2. Condemnation codes — animal didn't make it through processing.
    //      Pattern: any grade starting with 'X' (XML, XO, XTOX, future X-codes).
    //      No carcass weight, no health score required. Routes to IsCondemned=true.
    //
    //   3. Anything else — invalid; flagged for operator review.
    public static class GradeRules
    {
        // Production grades for cows. NP (Nicholas Prime) included because operators
        // manually enter it post-pull as a meaningful classification.
        public static readonly HashSet<string> CowGrades = new(StringComparer.OrdinalIgnoreCase)
        {
            "CN", "SH", "CT", "B1", "B2", "BR", "NP"
        };

        // Production grades for bulls. NP also valid here per operations spec.
        public static readonly HashSet<string> BullGrades = new(StringComparer.OrdinalIgnoreCase)
        {
            "BB", "LB", "UB", "FB", "NP"
        };

        // Pattern-based condemnation detection: any grade starting with 'X'.
        // This is forward-compatible — new X-codes added at the kill floor work
        // automatically without code changes.
        public static bool IsCondemnationCode(string? grade)
        {
            if (string.IsNullOrWhiteSpace(grade)) return false;
            var trimmed = grade.Trim();
            return trimmed.Length > 0 && (trimmed[0] == 'X' || trimmed[0] == 'x');
        }

        // Validates a grade against the animal's sex/type. Returns null if valid,
        // or a human-readable error message if invalid.
        //
        // Condemnation codes always pass (no allow-list check; condemned animals
        // bypass production grade validation entirely).
        //
        // When sex/type can't be classified as Cow or Bull, falls back to "valid
        // if grade is in either allow-list" — permissive for unknown classifications.
        public static string? ValidateGrade(string? sex, string? animalType, string? grade)
        {
            if (string.IsNullOrWhiteSpace(grade))
                return "Grade required";

            var g = grade.Trim();

            // Condemnation codes bypass allow-list checks
            if (IsCondemnationCode(g)) return null;

            var t = (animalType ?? string.Empty).Trim().ToUpperInvariant();
            var s = (sex ?? string.Empty).Trim().ToUpperInvariant();

            bool isCow  = t.Contains("COW")  || s == "F";
            bool isBull = t.Contains("BULL") || s == "B" || s == "M";

            if (isCow && CowGrades.Contains(g))   return null;
            if (isBull && BullGrades.Contains(g)) return null;

            if (isCow)
                return $"Grade {g} is not valid for Cow. Allowed: {string.Join(", ", CowGrades)}";
            if (isBull)
                return $"Grade {g} is not valid for Bull. Allowed: {string.Join(", ", BullGrades)}";

            // Sex/Type unclear — accept if grade is valid for either type
            if (CowGrades.Contains(g) || BullGrades.Contains(g)) return null;
            return $"Grade {g} is not in any allow-list (Sex={sex}, Type={animalType})";
        }
    }
}
