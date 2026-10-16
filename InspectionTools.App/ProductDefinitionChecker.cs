using InspectionTools.Inspection;
using InspectionTools.Products;

namespace InspectionTools.App;

/// <summary>製品の定義と、測定器の手順表の食い違いを、製品を開くときに見つける</summary>
public static class ProductDefinitionChecker {

    /// <summary>
    /// 定義の項目の動作 ID(actionId)が、すべて手順表にあるか確かめる。なければ、どの項目かを添えて ProductDefinitionException。
    /// 手順表にない ID は、検査中の動作(項目の移動、値の取得など)で例外になるので、開く時点で知らせてメニューに戻す
    /// </summary>
    public static void EnsureActionsExist(ProductDefinition definition, ProductProcedure procedure) {
        var unknown = definition.Items
            .Where(item => item.ActionId is not null && !procedure.Actions.ContainsKey(item.ActionId))
            .Select(item => $"{item.Name}(actionId: {item.ActionId})")
            .ToList();
        if (unknown.Count > 0) {
            throw new ProductDefinitionException($"{procedure.Product} の手順表にない動作 ID(actionId)が、定義にあります: {string.Join("、", unknown)}");
        }
    }
}
