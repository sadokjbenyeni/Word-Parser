using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.ModelBinding;
using System.Web.Http.OData;
using System.Web.Http.OData.Query;
using System.Web.Http.OData.Routing;
using Services.Services;
using Microsoft.Data.OData;

namespace Api.Controllers
{
    /*
    The WebApiConfig class may require additional changes to add a route for this controller. Merge these statements into the Register method of the WebApiConfig class as applicable. Note that OData URLs are case sensitive.

    using System.Web.Http.OData.Builder;
    using System.Web.Http.OData.Extensions;
    using Services.Services;
    ODataConventionModelBuilder builder = new ODataConventionModelBuilder();
    builder.EntitySet<FunctionalTestDocument>("FunctionalTestDocuments");
    config.Routes.MapODataServiceRoute("odata", "odata", builder.GetEdmModel());
    */
    public class FunctionalTestDocumentsController : ODataController
    {
        private static ODataValidationSettings _validationSettings = new ODataValidationSettings();

        // GET: odata/FunctionalTestDocuments
        public IHttpActionResult GetFunctionalTestDocuments(ODataQueryOptions<FunctionalTestDocument> queryOptions)
        {
            // validate the query.
            try
            {
                queryOptions.Validate(_validationSettings);
            }
            catch (ODataException ex)
            {
                return BadRequest(ex.Message);
            }

            // return Ok<IEnumerable<FunctionalTestDocument>>(functionalTestDocuments);
            return StatusCode(HttpStatusCode.NotImplemented);
        }

        // GET: odata/FunctionalTestDocuments(5)
        public IHttpActionResult GetFunctionalTestDocument([FromODataUri] int key, ODataQueryOptions<FunctionalTestDocument> queryOptions)
        {
            // validate the query.
            try
            {
                queryOptions.Validate(_validationSettings);
            }
            catch (ODataException ex)
            {
                return BadRequest(ex.Message);
            }

            // return Ok<FunctionalTestDocument>(functionalTestDocument);
            return StatusCode(HttpStatusCode.NotImplemented);
        }

        // PUT: odata/FunctionalTestDocuments(5)
        public IHttpActionResult Put([FromODataUri] int key, Delta<FunctionalTestDocument> delta)
        {
            Validate(delta.GetEntity());

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            // TODO: Get the entity here.

            // delta.Put(functionalTestDocument);

            // TODO: Save the patched entity.

            // return Updated(functionalTestDocument);
            return StatusCode(HttpStatusCode.NotImplemented);
        }

        // POST: odata/FunctionalTestDocuments
        public IHttpActionResult Post(FunctionalTestDocument functionalTestDocument)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            // TODO: Add create logic here.

            // return Created(functionalTestDocument);
            return StatusCode(HttpStatusCode.NotImplemented);
        }

        // PATCH: odata/FunctionalTestDocuments(5)
        [AcceptVerbs("PATCH", "MERGE")]
        public IHttpActionResult Patch([FromODataUri] int key, Delta<FunctionalTestDocument> delta)
        {
            Validate(delta.GetEntity());

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            // TODO: Get the entity here.

            // delta.Patch(functionalTestDocument);

            // TODO: Save the patched entity.

            // return Updated(functionalTestDocument);
            return StatusCode(HttpStatusCode.NotImplemented);
        }

        // DELETE: odata/FunctionalTestDocuments(5)
        public IHttpActionResult Delete([FromODataUri] int key)
        {
            // TODO: Add delete logic here.

            // return StatusCode(HttpStatusCode.NoContent);
            return StatusCode(HttpStatusCode.NotImplemented);
        }
    }
}
