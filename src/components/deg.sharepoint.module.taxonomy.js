shpUtility.factory('shpTaxonomy', ['$http', function ($http) {

    return {
        GetTermSetValues: getTermSetValues
    }

    function getTermSetValues(taxonomyGroup, termSetName, callback) {
        var context = SP.ClientContext.get_current();

        var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        var termStore = session.getDefaultSiteCollectionTermStore();
        var group = termStore.get_groups().getByName(taxonomyGroup);
        var termSet = group.get_termSets().getByName(termSetName);
        var terms = termSet.getAllTerms();

        context.load(terms);
        context.executeQueryAsync(
            function () {
                var values = [];
                var termEnumerator = terms.getEnumerator();
                while (termEnumerator.moveNext()) {
                    var currentTerm = termEnumerator.get_current();
                    values.push({ 'id': currentTerm.get_id(), 'name': currentTerm.get_name() });
                }
                if (callback) callback(values);
            },
            function (sender, args) {
                $log.log(args.get_message());
            }
        );
    }
}]);