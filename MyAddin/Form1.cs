using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace MyAddin
{
    public partial class Form1 : Form
    {
        private XmlDocument docXML = new XmlDocument();

        //this is the path where xml files are saved and loaded
        //the input and output file will be found in AppData/roaming folder
        string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string fileName = Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.ApplicationData), "input.xml");

        string fileName1 = Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.ApplicationData), "GenericFTS.xml");
        String name1 = "";

        // Xml tag for node, e.g. 'node' in case of <node></node>
        public const string XmlNodeTag = "node";

        // Xml attributes for node e.g. <node text="Sporadic Timing Failure"></node>
        public const string XmlNodeTextAtt = "text";

        public Form1()
        {
            try
            {
                InitializeComponent();
                //this will chk if xml exist and load that
                if (File.Exists(fileName))
                {
                    LoadXml(); //if xml exist then it will load the xml
                    treeView1.ExpandAll();
                }
                else
                {
                    GenericFTS(); //else generic FTS loaded
                    LoadXml();
                    treeView1.ExpandAll();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Invalid XML", "Error", MessageBoxButtons.OK);
                Console.WriteLine(e);
            }
        }

        public void GenericFTS()
        {
            try
            {
                treeView1.Nodes.Add("Failure");
                treeView1.Nodes[0].Nodes.Add("Sporadic Timing Failure");
                treeView1.Nodes[0].Nodes[0].Nodes.Add("Sporadic Duration Failure");
                treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes.Add("SporadicProlongation");
                treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[0].Nodes.Add("Permanent Prolongation");
                treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes.Add("Permanent Duration Failure");
                treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[1].Nodes.Add("Permanent Prolongation");
                treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[1].Nodes.Add("Permament Interupt");
            }
            catch (Exception)
            {
                MessageBox.Show("Not Valid GenericFTS", "Error", MessageBoxButtons.OK);
            }

        }

        public void LoadXml()
        {
            try
            {
                ReadXMLFile(treeView1, fileName);
            }
            catch (Exception)
            {
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
        }

        //this button contains functionalities for add button
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (treeView1.SelectedNode == null)
                {
                    MessageBox.Show("Please Select a Node to add new node.", "Error", MessageBoxButtons.OK);
                }

                else if (name1 == "")
                {
                    MessageBox.Show("Please Enter name of Node to add new node.", "Error", MessageBoxButtons.OK);
                }
                else
                {
                    TreeNode node = treeView1.SelectedNode;
                    node.Nodes.Add(name1);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Please Select a Node to add new node.", "Error", MessageBoxButtons.OK);
            }
            treeView1.SelectedNode = null;
            textBox1.Clear();
            return;
        }

        //here user give name for new node
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                name1 = textBox1.Text;  //get name from user for new node
            }
            catch (Exception)
            {
                MessageBox.Show("Not a Valid Name.", "Error", MessageBoxButtons.OK);
            }
        }

        //this button contain functionality to remove node
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                treeView1.SelectedNode.Remove();
            }
            catch (Exception)
            {
                MessageBox.Show("Please Select a Node to remove.", "Error", MessageBoxButtons.OK);
            }
            treeView1.SelectedNode = null;
            return;
        }

        //this button contain functionality to reset treeview
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                treeView1.Nodes.Clear();
                textBox1.Clear();
                GenericFTS(); //else generic FTS loaded
                treeView1.ExpandAll();
            }
            catch (Exception)
            {
                MessageBox.Show("Unable to reset.", "Error", MessageBoxButtons.OK);
            }
        }

        ////this button contain functionality to close addin
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Confirm?", "Close Application", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    this.Close();
                }
                else
                {
                    this.Activate();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Problem Closing the Application.", "Error", MessageBoxButtons.OK);
                treeView1.Nodes.Clear();
                GenericFTS(); //else generic FTS loaded
                treeView1.ExpandAll();
            }
        }

        //this button contain functionality to save xml files
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                TreeViewToXML(treeView1, fileName1);
                MessageBox.Show("XML created successfully.", "Done", MessageBoxButtons.OK);
            }
            catch (Exception)
            {
                MessageBox.Show("Problem Saving XML file, Try again", "Error", MessageBoxButtons.OK);
                textBox1.Clear();
                treeView1.Nodes.Clear();
                GenericFTS(); //else generic FTS loaded
                treeView1.ExpandAll();
            }
        }

        public void TreeViewToXML(TreeView treeView, string fileName)
        {
            XmlTextWriter textWriter = new XmlTextWriter(fileName, System.Text.Encoding.ASCII);
            // writing the xml declaration tag
            textWriter.WriteStartDocument();

            // writing the main tag that encloses all node tags
            textWriter.WriteStartElement("TreeView");

            // save the nodes, recursive method
            SaveNodes(treeView.Nodes, textWriter);
            textWriter.WriteEndElement();
            textWriter.Close();
        }

        private void SaveNodes(TreeNodeCollection nodesCollection, XmlTextWriter textWriter)
        {
            for (int i = 0; i < nodesCollection.Count; i++)
            {
                TreeNode node = nodesCollection[i];
                textWriter.WriteStartElement(XmlNodeTag);
                textWriter.WriteAttributeString(XmlNodeTextAtt, node.Text);

                if (node.Nodes.Count > 0)
                {
                    SaveNodes(node.Nodes, textWriter);
                }
                textWriter.WriteEndElement();
            }
        }

        public void ReadXMLFile(TreeView treeView, string fileName)
        {
            XmlTextReader reader = null;
            try
            {
                // disabling re-drawing of treeview till all nodes are added
                treeView.BeginUpdate();
                reader = new XmlTextReader(fileName);

                TreeNode parentNode = null;

                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (reader.Name == XmlNodeTag)
                        {
                            TreeNode newNode = new TreeNode();
                            bool isEmptyElement = reader.IsEmptyElement;

                            // loading node attributes
                            int attributeCount = reader.AttributeCount;
                            if (attributeCount > 0)
                            {
                                for (int i = 0; i < attributeCount; i++)
                                {
                                    reader.MoveToAttribute(i);
                                    SetAttributeValue(newNode, reader.Name, reader.Value);
                                }
                            }

                            // add new node to Parent Node or TreeView
                            if (parentNode != null)
                                parentNode.Nodes.Add(newNode);
                            else
                                treeView.Nodes.Add(newNode);

                            // making current node 'ParentNode' if its not empty
                            if (!isEmptyElement)
                            {
                                parentNode = newNode;
                            }
                        }
                    }
                    // moving up to in TreeView if end tag is encountered
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        if (reader.Name == XmlNodeTag)
                        {
                            parentNode = parentNode.Parent;
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.XmlDeclaration)
                    { //Ignore Xml Declaration                    
                    }
                    else if (reader.NodeType == XmlNodeType.None)
                    {
                        return;
                    }
                    else if (reader.NodeType == XmlNodeType.Text)
                    {
                        parentNode.Nodes.Add(reader.Value);
                    }
                }
            }
            finally
            {
                // enabling redrawing of treeview after all nodes are added
                treeView.EndUpdate();
                reader.Close();
            }
        }

        public void SetAttributeValue(TreeNode node, string propertyName, string value)
        {
            if (propertyName == XmlNodeTextAtt)
            {
                node.Text = value;
            }
        }
    }
}



