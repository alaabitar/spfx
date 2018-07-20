import * as React from 'react';
import styles from './Taxonomynavigation.module.scss';
import { ITaxonomynavigationProps } from './ITaxonomynavigationProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Utils } from "../utils/Utils";

export interface ITaxonomynavigation{
  items:Array<any>,
  loadingScripts: boolean,
  errors: Array<any>
}

export default class Taxonomynavigation extends React.Component<ITaxonomynavigationProps, ITaxonomynavigation> {
  
  public constructor(props: ITaxonomynavigationProps){
    super(props);
    this.state = {
      items: [],
      loadingScripts: true,
      errors: []
    };
  }
  
  public componentDidMount(){
    this._loadSPJSOMScripts();
  }

  private _loadSPJSOMScripts() {
    const siteColUrl = Utils.getSiteCollectionUrl();
    try {
      SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/init.js', {
        globalExportsName: '$_global_init'
      })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/MicrosoftAjax.js', {
            globalExportsName: 'Sys'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.Runtime.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.taxonomy.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): void => {
          this.setState({ loadingScripts: false });
          const context: SP.ClientContext = new SP.ClientContext(this.props.siteUrl);
          let taxSession =  SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
          let termStore  = taxSession.getDefaultSiteCollectionTermStore();
          let termGroups = termStore.get_groups();
          let termGroup = termGroups.getByName("UNAIDS");
          let termSets = termGroup.get_termSets();
          this.loadTermStore(termSets, context);
        })
        .catch((reason: any) => {

        });
    } catch (error) {

    }
  }

  private loadTermStore(termSets: SP.Taxonomy.TermSetCollection,spContext:SP.ClientContext ){

    var reactHandler = this;

    let termSet = termSets.getByName("NAVIGATION");
    let terms = termSet.getAllTerms();

    spContext.load(terms, 'Include(Name, Parent, IsRoot,Id,PathOfTerm)');

    spContext.load(terms);
    spContext.load(termSet);
    let termStore:any[]=[];
    let childTerm:any[]=[];
    spContext.executeQueryAsync(function () {

      var termsEnum = terms.getEnumerator();


      while (termsEnum.moveNext()) {

        var spTerm = termsEnum.get_current();
        termStore.push({label:spTerm.get_name(),value:spTerm.get_name(), id:spTerm.get_id(), pathOfTerm:spTerm.get_pathOfTerm(),pathOfParentTerm:  spTerm.get_isRoot()?"":spTerm.get_parent().get_pathOfTerm()});

      }

      window['termStore']= termStore;

      let currentTermUrl = document.location.href.replace(document.location.search,'').replace(spContext.get_url()+'/','').replace(/\//g,';').replace(/-/g,' ').toLowerCase();
      termStore.filter((e) => e.pathOfParentTerm.toLowerCase() === currentTermUrl).forEach( term =>{
        childTerm.push({Url:reactHandler.props.siteUrl + "/" + term.pathOfTerm.replace(/;/g,'/').replace(/ /g,'-'),Description:term.value});
      });

      reactHandler.setState({
        items: childTerm
      });

    });
  }

  public render(): React.ReactElement<ITaxonomynavigationProps> {
    return (
      <div className="QuickLinks">
        <h1 className="QuickLinks h1">{this.props.description}</h1>
        <div className="QuickLinks div">
          <table>
            {this.state.items.map(function(item,key){
                    return (
                        <tr><td><a className="serialQuickLinks a" href={item.Url}>{item.Description}</a></td></tr>
                      );
                  })}
          </table>
        </div>
      </div>
    );
  }
}
